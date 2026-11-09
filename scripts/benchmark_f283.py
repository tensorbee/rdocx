#!/usr/bin/env python3
"""Exact Issue445 workloads with authenticated final-context admission.
Preparation self-tests never import or execute the product or measure time.
"""
import argparse
import ast
import contextlib
import hashlib
import importlib
import importlib.metadata
import io
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import sysconfig
import zipfile
import xml.etree.ElementTree as ET

ORIGINAL_CASE1 = '"""Case 1: Cell.text, Cell.vertical_alignment and Cell.shading cost grows with the body.\n\nBuilds its own input: a body of N paragraphs, then one 20 x 5 table whose 100 cells are set.\nPASS when each setter costs under 1 ms per call at 2000 paragraphs and at most 1.5 times its empty-body cost\n(plus 0.05 ms, as timings that small are noise).\n"""\nimport platform\nfrom importlib.metadata import version\nimport sys\nimport time\n\nimport rdocx\n\nprint(f"rdocx {version(\'rdocx\')}, Python {platform.python_version()}, {platform.system()} {platform.machine()}")\nSETTERS = {\n    "Cell.text": lambda cell: setattr(cell, "text", "12.50"),\n    "Cell.vertical_alignment": lambda cell: setattr(\n        cell, "vertical_alignment", rdocx.WD_CELL_VERTICAL_ALIGNMENT.CENTER\n    ),\n    "Cell.shading": lambda cell: setattr(cell, "shading", "D9E2F3"),\n}\ncosts = {name: {} for name in SETTERS}\nfor body in (0, 500, 2000):\n    doc = rdocx.Document()\n    for index in range(body):\n        doc.add_paragraph(f"Body paragraph {index} with some text in it to look like a report.")\n    table = doc.add_table(20, 5)\n    cells = [table.cell(row, column) for row in range(20) for column in range(5)]\n    for name, setter in SETTERS.items():\n        started = time.perf_counter()\n        for cell in cells:\n            setter(cell)\n        costs[name][body] = 1000 * (time.perf_counter() - started) / len(cells)\nfailed = False\nfor name, cost in costs.items():\n    ok = cost[2000] < 1.0 and cost[2000] <= 1.5 * cost[0] + 0.05\n    failed |= not ok\n    print(f"{\'PASS\' if ok else \'FAIL\'} {name}: "\n          + ", ".join(f"{cost[body]:.2f} ms at {body}" for body in cost)\n          + f" paragraphs (x{cost[2000] / cost[0]:.1f})")\nsys.exit(1 if failed else 0)\n'
ORIGINAL_CASE2 = '"""Case 2: Document.open + try_replace_text + save in Python costs about twice `rdocx replace`.\n\nBuilds its own input: 4000 paragraphs of three runs each, "water" in every fourth one.\nRun with RDOCX_CLI set to the rdocx CLI of the same build.\nPASS when the Python total (open, replace, save) is at most 1.2 times the CLI\'s time, process start included.\n"""\nimport os\nimport platform\nfrom importlib.metadata import version\nimport subprocess\nimport sys\nimport tempfile\nimport time\n\nimport rdocx\n\nprint(f"rdocx {version(\'rdocx\')}, Python {platform.python_version()}, {platform.system()} {platform.machine()}")\ncli = os.environ.get("RDOCX_CLI", "rdocx")\nwith tempfile.TemporaryDirectory() as folder:\n    doc = rdocx.Document()\n    for index in range(4000):\n        paragraph = doc.add_paragraph()\n        paragraph.add_run(f"Section {index}: ").font.bold = True\n        paragraph.add_run("the plant treats " + ("water" if index % 4 == 0 else "air") + " for the town, ")\n        paragraph.add_run("as the report explains.").font.italic = True\n    source = os.path.join(folder, "in.docx")\n    doc.save(source)\n    python, command = [], []\n    for trial in range(3):\n        t0 = time.perf_counter()\n        doc = rdocx.Document.open(source)\n        t1 = time.perf_counter()\n        count = doc.try_replace_text("water", "WATER")\n        t2 = time.perf_counter()\n        doc.save(os.path.join(folder, f"py{trial}.docx"))\n        t3 = time.perf_counter()\n        python.append((t3 - t0, t1 - t0, t2 - t1, t3 - t2))\n        started = time.perf_counter()\n        subprocess.run([cli, "replace", source, "-p", "water", "-v", "WATER",\n                        "-o", os.path.join(folder, f"cli{trial}.docx")], check=True, capture_output=True)\n        command.append(time.perf_counter() - started)\ntotal, opened, replaced, saved = min(python)\nbest_cli = min(command)\nok = total <= 1.2 * best_cli\nprint(f"{\'PASS\' if ok else \'FAIL\'} {count} replaced: Python open {opened:.2f} s + replace {replaced:.2f} s "\n      f"+ save {saved:.2f} s = {total:.2f} s; CLI {best_cli:.2f} s (x{total / best_cli:.1f})")\nsys.exit(0 if ok else 1)\n'
PINS = {'case1': '81d82e4620df39832a027d55cb475c6863e35e729da61646c2c7eb970a1e983a',
        'case2': '4e13939a43630495d9397d85ae985cdf2a597a4488c4ac822033d6c900ea0841'}


def sha(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def require(condition, message):
    if not condition:
        raise RuntimeError(message)


def atomic_bytes(path, payload):
    path = Path(path)
    require(not path.exists(), 'Refuse report overwrite: ' + str(path))
    temp = path.with_suffix(path.suffix + '.partial')
    created = False
    try:
        stream = temp.open('xb')
        created = True
        with stream:
            stream.write(payload)
        os.link(temp, path)
    finally:
        if created:
            temp.unlink(missing_ok=True)


def atomic_json(path, value):
    atomic_bytes(path, (json.dumps(value, indent=2, allow_nan=False) + '\n').encode('utf-8'))


def flat_score(empty, large):
    return large < 1.0 and large <= 1.5 * empty + 0.05


def pair_score(python, command):
    require(len(python) == len(command) == 3, 'Need all three paired trials')
    best_python, best_cli = min(python), min(command)
    return {'python_tuple': best_python, 'cli_seconds': best_cli,
            'pass': best_python[0] <= 1.2 * best_cli}


def host_facts():
    return {'system': platform.system(), 'machine': platform.machine(),
            'implementation': platform.python_implementation(),
            'version': list(sys.version_info[:3]), 'sysconfig_platform': sysconfig.get_platform(),
            'executable': str(Path(sys.executable).resolve()), 'executable_sha256': sha(sys.executable)}


def admit_host(facts):
    require(facts['system'] == 'Linux' and facts['machine'] == 'x86_64'
            and facts['implementation'] == 'CPython' and facts['version'][:2] == [3, 12]
            and facts['sysconfig_platform'] == 'linux-x86_64',
            'Requires native Linux x86_64 CPython3.12; no timing permitted')


def checked_file(record):
    path = Path(record['path'])
    require(path.is_absolute() and path.is_file(), 'Missing absolute file: ' + str(path))
    require(sha(path) == record['sha256'], 'File identity mismatch: ' + str(path))
    return path


def check_no_compilers():
    output = subprocess.check_output(['ps', '-eo', 'comm='], text=True)
    forbidden = {'cargo', 'rustc', 'cc', 'clang', 'clang++', 'gcc', 'g++', 'maturin', 'ninja'}
    require(not any(Path(line.strip()).name in forbidden for line in output.splitlines()),
            'Compiler/build process present; exclusive timing host required')
    return {'process_snapshot_sha256': hashlib.sha256(output.encode()).hexdigest(),
            'no_named_compiler_observed': True}


def authenticate(proof, artifact, import_product):
    facts = host_facts()
    admit_host(facts)
    require(proof['schema'] == 'f283-root-approved-final-input-v1', 'Unknown proof schema')
    require(proof['reviewed_scope'] == 'F283 final Linux release gates'
            and proof['native_execution'] is True and proof['exclusive_host_reservation'],
            'Missing reviewed native/exclusive-host producer evidence')
    require(proof['runner_sha256'] == sha(__file__), 'Runner source mismatch')
    require(proof['interpreter_sha256'] == facts['executable_sha256'], 'Interpreter mismatch')
    repo = Path(proof['source_root'])
    require(repo.is_absolute(), 'Source root must be absolute')
    head = subprocess.check_output(['git', '-C', str(repo), 'rev-parse', 'HEAD'], text=True).strip()
    require(head == proof['source_commit'], 'Source HEAD differs from reviewed build source')
    require(not subprocess.check_output(['git', '-C', str(repo), 'status', '--porcelain']),
            'Final source must be clean')
    manifest_path = checked_file(proof['source_manifest'])
    inputs = json.loads(manifest_path.read_text())
    require(isinstance(inputs, dict) and inputs, 'Empty source input manifest')
    for name, digest in inputs.items():
        require(not Path(name).is_absolute() and '..' not in Path(name).parts, 'Unsafe source path')
        require(sha(repo / name) == digest, 'Source input mismatch: ' + name)
    wheel, cli = checked_file(proof['wheel']), checked_file(proof['cli'])
    require(cli.read_bytes()[:5] == b'\x7fELF\x02' and int.from_bytes(cli.read_bytes()[18:20], 'little') == 62,
            'CLI must be native x86_64 ELF64')
    for kind in ['wheel', 'cli']:
        require(proof['build_mode_source_attestation_accepted'] is True, 'Build-mode/source attestation still pending Root review')
        require(proof['build_receipts'][kind]['sha256'] == proof['approved_build_attestation_sha256'][kind],
                'Producer receipt not independently pinned by Root-approved input')
        receipt = json.loads(checked_file(proof['build_receipts'][kind]).read_text())
        expected_build = proof['approved_build_configurations'][kind]
        require(all(receipt[key] == expected_build[key] for key in ['argv', 'environment', 'profile', 'toolchain']),
                'Build configuration differs from Root-approved actual producer evidence')
        require(receipt['kind'] == kind and receipt['source_commit'] == head
                and receipt['source_manifest_sha256'] == proof['source_manifest']['sha256']
                and receipt['artifact_sha256'] == proof[kind]['sha256']
                and receipt['target'] == 'x86_64-unknown-linux-gnu'
                and receipt['child_exit'] == 0, 'Build receipt source/artifact mismatch: ' + kind)
        require('--release' in receipt['argv'] and '--locked' in receipt['argv']
                and receipt['profile']['name'] == 'release'
                and receipt['profile']['debug_assertions'] is False
                and receipt['profile']['overflow_checks'] is False
                and str(receipt['profile']['opt_level']) == '3'
                and receipt['toolchain'] and receipt['environment'],
                'Release configuration not proved: ' + kind)
    with zipfile.ZipFile(wheel) as archive:
        metadata = [n for n in archive.namelist() if n.endswith('.dist-info/WHEEL')]
        require(len(metadata) == 1, 'Ambiguous wheel metadata')
        tags = [line[5:] for line in archive.read(metadata[0]).decode().splitlines() if line.startswith('Tag: ')]
        require(any(t.startswith('cp39-abi3-') and t.endswith('x86_64') and 'linux' in t for t in tags),
                'Wrong release wheel ABI/platform')
        distribution = importlib.metadata.distribution('rdocx')
        payload = [n for n in archive.namelist() if n.startswith('rdocx/') and not n.endswith('/')]
        require(payload and any(n.endswith('.so') for n in payload), 'Missing wheel package payload')
        for name in payload:
            installed = Path(distribution.locate_file(name)).resolve()
            require(installed.read_bytes() == archive.read(name), 'Installed wheel payload mismatch: ' + name)
        extension = [n for n in payload if n.endswith('.so')]
        require(len(extension) == 1, 'Ambiguous extension')
        extension_path = Path(distribution.locate_file(extension[0])).resolve()
    # Import happens only after host, reviewed build proof and installed bytes qualify.
    product = importlib.import_module('rdocx') if import_product else sys.modules['rdocx']
    require(Path(product.__file__).resolve() == Path(distribution.locate_file('rdocx/__init__.py')).resolve(),
            'Public import shadows installed package')
    require(any(getattr(module, '__file__', None) and Path(module.__file__).resolve() == extension_path
                for module in list(sys.modules.values())), 'Loaded extension differs from authenticated wheel')
    artifact.update({'host': facts, 'source_commit': head, 'wheel_sha256': sha(wheel),
                     'cli_sha256': sha(cli), 'extension_sha256': sha(extension_path),
                     'rdocx_version': distribution.version, 'wheel_tags': tags,
                     'build_proof_boundary': 'Caller-pinned reviewed producer receipts, not filename inference',
                     'compiler_check': check_no_compilers()})
    return product, str(cli)


class Instrument(ast.NodeTransformer):
    def visit_BinOp(self, node):
        self.generic_visit(node)
        if (isinstance(node.left, ast.Call) and ast.unparse(node.left) == 'time.perf_counter()'
                and isinstance(node.op, ast.Sub) and isinstance(node.right, ast.Name)
                and node.right.id == 'started'):
            return ast.NamedExpr(target=ast.Name(id='_f283_elapsed', ctx=ast.Store()), value=node)
        return node

    def visit_For(self, node):
        self.generic_visit(node)
        if isinstance(node.target, ast.Name) and node.target.id == 'body':
            node.body.append(ast.parse('_f283_correctness(body, cells)').body[0])
        return node

    def visit_Assign(self, node):
        self.generic_visit(node)
        if ast.unparse(node.targets[0]) == 'costs[name][body]':
            return [node, ast.parse('_f283_sample(body, name, _f283_elapsed, costs[name][body])').body[0]]
        return node

    def visit_Expr(self, node):
        self.generic_visit(node)
        if not isinstance(node.value, ast.Call):
            return node
        call = node.value
        callee = ast.unparse(call.func)
        if callee == 'sys.exit':
            return ast.Assign(targets=[ast.Name(id='_f283_gate_exit', ctx=ast.Store())], value=call.args[0])
        if callee == 'doc.save' and ast.unparse(call.args[0]) == 'source':
            return [node, ast.parse('_f283_source(source)').body[0]]
        if callee == 'command.append':
            return [node, ast.parse('_f283_trial(trial, count, python[-1], command[-1], (t0, t1, t2, t3), folder)').body[0]]
        return node


class Strip(ast.NodeTransformer):
    def visit_Expr(self, node):
        if isinstance(node.value, ast.Call) and ast.unparse(node.value.func).startswith('_f283_'):
            return None
        return self.generic_visit(node)

    def visit_NamedExpr(self, node):
        require(node.target.id.startswith('_f283_'), 'Unexpected inserted expression')
        return self.visit(node.value)

    def visit_Assign(self, node):
        if isinstance(node.targets[0], ast.Name) and node.targets[0].id == '_f283_gate_exit':
            return ast.Expr(value=ast.Call(func=ast.Attribute(value=ast.Name(id='sys', ctx=ast.Load()), attr='exit', ctx=ast.Load()), args=[node.value], keywords=[]))
        return self.generic_visit(node)


def transformed(source):
    original = ast.parse(source)
    instrumented = ast.fix_missing_locations(Instrument().visit(ast.parse(source)))
    stripped = ast.fix_missing_locations(Strip().visit(ast.parse(ast.unparse(instrumented))))
    require(ast.dump(original) == ast.dump(stripped), 'Instrumentation changed original sampler/workload AST')
    return instrumented


def self_test():
    require('rdocx' not in sys.modules, 'Self-tests must not import product')
    checks = {}
    checks['flat_inclusive_relative'] = flat_score(0.1, 1.5 * 0.1 + 0.05)
    checks['flat_above_relative_refused'] = not flat_score(0.1, 0.20000001)
    checks['flat_absolute_one_refused'] = not flat_score(1, 1)
    checks['case2_inclusive'] = pair_score([(1.2, .3, .4, .5)] * 3, [1, 2, 3])['pass']
    checks['case2_above_refused'] = not pair_score([(1.20001, .3, .4, .50001)] * 3, [1, 2, 3])['pass']
    tuples = [(6, 1, 4, 1), (5, 3, 1, 1), (5, 2, 2, 1)]
    checks['lexicographic_whole_tuple'] = pair_score(tuples, [7, 4, 8])['python_tuple'] == (5, 2, 2, 1)
    checks['cli_independent_minimum'] = pair_score(tuples, [7, 4, 8])['cli_seconds'] == 4
    try:
        pair_score(tuples[:2], [1, 2])
        checks['partial_trials_refused'] = False
    except RuntimeError:
        checks['partial_trials_refused'] = True
    try:
        admit_host({'system': 'Darwin', 'machine': 'arm64', 'implementation': 'CPython', 'version': [3, 12], 'sysconfig_platform': 'macosx'})
        checks['wrong_host_refused'] = False
    except RuntimeError:
        checks['wrong_host_refused'] = True
    for name, source in [('case1', ORIGINAL_CASE1), ('case2', ORIGINAL_CASE2)]:
        checks[name + '_pinned_bytes'] = hashlib.sha256(source.encode()).hexdigest() == PINS[name]
        transformed(source)
        checks[name + '_strip_ast_equivalent'] = True
    checks.update(source_shape_self_test())
    require(all(checks.values()), 'Synthetic self-test failed')
    return {'scope': 'Synthetic data and AST only. No product import, execution or timing.',
            'preparation_interpreter': sys.version, 'product_imported': 'rdocx' in sys.modules, 'checks': checks}


def verify_case2_source(package):
    """Independent untimed OOXML geometry check; never normalizes the input."""
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    with zipfile.ZipFile(package) as archive:
        require(archive.namelist().count('word/document.xml') == 1, 'Ambiguous source document part')
        document = ET.fromstring(archive.read('word/document.xml'))
    require(document.tag == w + 'document', 'Wrong source document root')
    bodies = document.findall(w + 'body')
    require(len(bodies) == 1, 'Source must have one direct body')
    children = list(bodies[0])
    paragraphs = [child for child in children if child.tag == w + 'p']
    require(len(paragraphs) == 4000, 'Source must have exactly 4000 direct paragraphs')
    require(all(child.tag == w + 'p' for child in children)
            or len(children) == 4001 and children[-1].tag == w + 'sectPr'
            and all(child.tag == w + 'p' for child in children[:-1]),
            'Unexpected source body placement')

    def enabled(run, property_name):
        properties = run.findall(w + 'rPr')
        require(len(properties) <= 1, 'Duplicate source run properties')
        values = properties[0].findall(w + property_name) if properties else []
        return len(values) == 1 and values[0].get(w + 'val', 'true') in {'true', '1', 'on'}

    water = 0
    for index, paragraph in enumerate(paragraphs):
        require(all(child.tag in {w + 'pPr', w + 'r'} for child in paragraph), 'Non-direct source run placement')
        runs = paragraph.findall(w + 'r')
        require(len(runs) == 3, 'Source paragraph must have three direct runs')
        expected = [f'Section {index}: ',
                    'the plant treats ' + ('water' if index % 4 == 0 else 'air') + ' for the town, ',
                    'as the report explains.']
        observed_texts = []
        for run, text in zip(runs, expected):
            require(all(child.tag in {w + 'rPr', w + 't'} for child in run), 'Unexpected source run content')
            actual = ''.join(child.text or '' for child in run.findall(w + 't'))
            observed_texts.append(actual)
            require(actual == text, 'Source run text or placement differs at paragraph ' + str(index))
        require(enabled(runs[0], 'b'), 'Source first run is not bold')
        require(enabled(runs[2], 'i'), 'Source third run is not italic')
        water += 'water' in observed_texts[1]
    require(water == 1000, 'Source must have exactly 1000 water-bearing middle runs')
    return {'direct_paragraphs': 4000, 'direct_runs_per_paragraph': 3,
            'bold_first_runs': 4000, 'italic_third_runs': 4000,
            'water_bearing_middle_runs': water, 'all_expected_text_and_placement': True}


def source_shape_self_test():
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    document = ET.Element(w + 'document')
    body = ET.SubElement(document, w + 'body')
    for index in range(4000):
        paragraph = ET.SubElement(body, w + 'p')
        for position, text in enumerate([f'Section {index}: ', 'the plant treats ' + ('water' if index % 4 == 0 else 'air') + ' for the town, ', 'as the report explains.']):
            run = ET.SubElement(paragraph, w + 'r')
            if position in [0, 2]:
                properties = ET.SubElement(run, w + 'rPr')
                ET.SubElement(properties, w + ('b' if position == 0 else 'i'))
            ET.SubElement(run, w + 't').text = text
    baseline = ET.tostring(document)

    def package(xml):
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w') as archive:
            archive.writestr('word/document.xml', xml)
        stream.seek(0)
        return stream

    checks = {'source_shape_valid': verify_case2_source(package(baseline))['water_bearing_middle_runs'] == 1000}
    for fault in ['run_count', 'bold_style', 'italic_style', 'text_placement', 'paragraph_count']:
        changed = ET.fromstring(baseline)
        changed_body = changed.find(w + 'body')
        paragraph = changed_body[0]
        if fault == 'run_count':
            paragraph.remove(paragraph[-1])
        elif fault == 'bold_style':
            paragraph[0].find(w + 'rPr').find(w + 'b').set(w + 'val', '0')
        elif fault == 'italic_style':
            paragraph[2].find(w + 'rPr').find(w + 'i').set(w + 'val', 'false')
        elif fault == 'text_placement':
            left, right = paragraph[1].find(w + 't'), paragraph[2].find(w + 't')
            left.text, right.text = right.text, left.text
        else:
            changed_body.remove(changed_body[-1])
        try:
            verify_case2_source(package(ET.tostring(changed)))
            checks['source_shape_' + fault + '_refused'] = False
        except RuntimeError:
            checks['source_shape_' + fault + '_refused'] = True
    return checks


def plain_docx(path):
    with zipfile.ZipFile(path) as archive:
        xml = ET.fromstring(archive.read('word/document.xml'))
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    return [''.join(t.text or '' for t in p.iter(ns + 't'))
            for p in xml.find(ns + 'body') if p.tag == ns + 'p']


def execute(args):
    out = Path(args.output)
    out.mkdir(parents=False, exist_ok=False)
    report = {'status': 'incomplete', 'complete': False, 'samples': [], 'case1_correctness': [], 'trials': []}
    exit_code = 2
    namespaces = []
    try:
        proof_path = checked_file({'path': str(Path(args.proof).resolve()), 'sha256': args.proof_sha256})
        proof = json.loads(proof_path.read_text())
        report['proof_sha256'] = args.proof_sha256
        product, cli = authenticate(proof, report, True)
        os.environ['RDOCX_CLI'] = cli

        def sample(body, name, elapsed, cost):
            report['samples'].append({'body': body, 'setter': name, 'calls': 100, 'elapsed_seconds': elapsed, 'ms_per_call': cost})

        def correctness(body, cells):
            values = [[c.text, c.vertical_alignment == product.WD_CELL_VERTICAL_ALIGNMENT.CENTER, c.shading] for c in cells]
            require(len(values) == 100 and all(v == ['12.50', True, 'D9E2F3'] for v in values), 'Cell correctness failure')
            report['case1_correctness'].append({'body': body, 'values': values})

        def source(path):
            before = sha(path)
            report['case2_input_geometry'] = verify_case2_source(path)
            require(sha(path) == before, 'Source geometry check changed input')
            shutil.copyfile(path, out / 'case2-input.docx')
            report['input_sha256'] = before

        def trial(index, count, python_tuple, command_seconds, clocks, folder):
            record = {'trial': index, 'count': count, 'python_tuple_seconds': python_tuple,
                      'python_clock_readings': clocks, 'cli_seconds': command_seconds,
                      'cli_exit': 0, 'cli_exit_basis': 'Original subprocess.run(check=True) returned successfully'}
            for kind in ['py', 'cli']:
                destination = out / f'{kind}{index}.docx'
                shutil.copyfile(Path(folder) / destination.name, destination)
                record[kind + '_output_sha256'] = sha(destination)
            report['trials'].append(record)

        for number, original in [(args.case, ORIGINAL_CASE1 if args.case == 1 else ORIGINAL_CASE2)]:
            report[f'case{number}_pre_compiler_check'] = check_no_compilers()
            namespace = {'__name__': '__f283_original__', '_f283_sample': sample, '_f283_correctness': correctness,
                         '_f283_source': source, '_f283_trial': trial}
            namespaces.append(namespace)
            with (out / f'case{number}.stdout').open('x') as stdout:
                with contextlib.redirect_stdout(stdout):
                    exec(compile(transformed(original), f'<pinned-issue445-case{number}>', 'exec'), namespace)
            report[f'case{number}_gate_exit'] = namespace['_f283_gate_exit']
            report[f'case{number}_post_compiler_check'] = check_no_compilers()
        if args.case == 1:
            costs = namespaces[0]['costs']
            report['case1_thresholds'] = {name: {'costs_ms': cost, 'absolute_limit_strict_ms': 1.0,
                                              'relative_limit_inclusive_ms': 1.5 * cost[0] + .05,
                                              'pass': flat_score(cost[0], cost[2000])} for name, cost in costs.items()}
            require(len(report['samples']) == 9 and len(report['case1_correctness']) == 3, 'Missing cell samples')
            passed = all(v['pass'] for v in report['case1_thresholds'].values())
        else:
            report['case2_threshold'] = pair_score(namespaces[0]['python'], namespaces[0]['command'])
            require(len(report['trials']) == 3, 'Missing paired trials')
            require(all(t['count'] == 1000 and t['cli_exit'] == 0 for t in report['trials']), 'Missing replacement count or CLI failure')
            expected = [f'Section {i}: the plant treats ' + ('WATER' if i % 4 == 0 else 'air') + ' for the town, as the report explains.' for i in range(4000)]
            for index in range(3):
                for kind in ['py', 'cli']:
                    require(plain_docx(out / f'{kind}{index}.docx') == expected, 'Saved output text mismatch')
            require(sha(out / 'case2-input.docx') == report['input_sha256'], 'Source copy changed')
            passed = report['case2_threshold']['pass']
        authenticate(proof, {}, False)
        report['proof_unchanged'] = sha(proof_path) == args.proof_sha256
        require(report['proof_unchanged'], 'Proof changed during timing')
        require(passed == (report[f'case{args.case}_gate_exit'] == 0), 'Original/scorer disagreement')
        report.update({'complete': True, 'status': 'pass' if passed else 'threshold_failure'})
        exit_code = 0 if passed else 1
    except Exception as error:
        report['partial_original_state'] = [{key: namespace[key] for key in ['costs', 'python', 'command', 'trial', 'count', 't0', 't1', 't2', 't3'] if key in namespace} for namespace in namespaces]
        report['failure'] = {'type': type(error).__name__, 'message': str(error)}
        if isinstance(error, subprocess.CalledProcessError):
            report['failed_child'] = {'argv': error.cmd, 'exit': error.returncode}
            for label, value in [('stdout', error.stdout), ('stderr', error.stderr)]:
                if value is not None:
                    (out / ('failed-cli.' + label)).write_bytes(value if isinstance(value, bytes) else value.encode())
    finally:
        report['exit_code'] = exit_code
        atomic_json(out / 'report.json', report)
    return exit_code


def run_both(args):
    out = Path(args.output)
    out.mkdir(parents=False, exist_ok=False)
    summary = {'complete': False, 'status': 'incomplete', 'workers': []}
    exit_code = 2
    try:
        self_test()
        admit_host(host_facts())
        checked_file({'path': str(Path(args.proof).resolve()), 'sha256': args.proof_sha256})
        for number in [1, 2]:
            destination = out / f'case{number}'
            argv = [sys.executable, '-I', '-B', str(Path(__file__).resolve()), '--proof', args.proof,
                    '--proof-sha256', args.proof_sha256, '--output', str(destination), '--case', str(number)]
            with (out / f'worker{number}.stdout').open('x') as stdout, (out / f'worker{number}.stderr').open('x') as stderr:
                child = subprocess.run(argv, stdout=stdout, stderr=stderr, check=False)
            worker = {'argv': argv, 'child_exit': child.returncode}
            summary['workers'].append(worker)
            result = json.loads((destination / 'report.json').read_text())
            worker.update({'report_sha256': sha(destination / 'report.json'), 'result': result})
            require(child.returncode in [0, 1] and result['complete'] and result['exit_code'] == child.returncode, 'Incomplete or failed worker')
        passed = all(worker['child_exit'] == 0 for worker in summary['workers'])
        summary.update({'complete': True, 'status': 'pass' if passed else 'threshold_failure'})
        exit_code = 0 if passed else 1
    except Exception as error:
        summary['failure'] = {'type': type(error).__name__, 'message': str(error)}
    finally:
        summary['exit_code'] = exit_code
        atomic_json(out / 'report.json', summary)
    return exit_code


def prepare_dispatch(args):
    """Package an already Root-approved transfer into a local dispatch request.

    Validation-only preparation of existing approved data. It never contacts a
    service, approves a producer, starts a benchmark or imports the product.
    """
    output = Path(args.output)
    require(output.is_absolute(), 'Dispatch request output must be an absolute path')
    require(re.fullmatch('[0-9a-f]{64}', args.transfer_sha256),
            'Transfer SHA-256 must be lowercase hexadecimal')
    transfer_path = Path(args.transfer).resolve()
    require(transfer_path.is_absolute() and transfer_path.is_file(),
            'Missing absolute file: ' + str(transfer_path))
    transfer_raw = transfer_path.read_bytes()
    require(hashlib.sha256(transfer_raw).hexdigest() == args.transfer_sha256,
            'File identity mismatch: ' + str(transfer_path))
    transfer_text = transfer_raw.decode('utf-8')
    try:
        transfer = json.loads(transfer_text)
    except json.JSONDecodeError as error:
        raise RuntimeError('Transfer is not valid JSON: ' + str(error)) from None
    require(isinstance(transfer, dict) and transfer.get('schema') == 'f283-root-transfer-v1',
            'Unknown transfer schema')
    for key in ['producer_run_id', 'wheel_artifact_id', 'cli_artifact_id']:
        value = transfer.get(key)
        require(not isinstance(value, bool) and str(value).isdigit() and int(value) > 0,
                'Nonpositive or noninteger producer/artifact ID: ' + key)
    proof_text = transfer.get('proof_json')
    require(isinstance(proof_text, str), 'Missing transfer proof text')
    require(hashlib.sha256(proof_text.encode('utf-8')).hexdigest() == transfer.get('proof_sha256'),
            'Proof JSON differs from the transfer-pinned SHA-256')
    try:
        proof = json.loads(proof_text)
    except json.JSONDecodeError as error:
        raise RuntimeError('Proof is not valid JSON: ' + str(error)) from None
    require(isinstance(proof, dict) and proof.get('schema') == 'f283-root-approved-final-input-v1',
            'Unknown proof schema')
    source_sha = proof.get('source_commit')
    require(isinstance(source_sha, str) and re.fullmatch('[0-9a-f]{40}', source_sha),
            'Proof source_commit must be lowercase 40-hex')
    require(proof.get('build_mode_source_attestation_accepted') is True,
            'Build-mode/source attestation still pending Root review')
    receipts = proof.get('build_receipts')
    require(isinstance(receipts, dict), 'Missing proof build receipts')
    for key, record in [('source_inputs_json', proof.get('source_manifest')),
                        ('wheel_build_receipt_json', receipts.get('wheel')),
                        ('cli_build_receipt_json', receipts.get('cli'))]:
        require(isinstance(record, dict) and isinstance(record.get('sha256'), str),
                'Incomplete nested proof record: ' + key)
        payload = transfer.get(key)
        require(isinstance(payload, str), 'Missing transfer receipt text: ' + key)
        require(hashlib.sha256(payload.encode('utf-8')).hexdigest() == record['sha256'],
                'Receipt identity differs from nested proof: ' + key)
    require(args.source_ref.startswith('refs/heads/') and len(args.source_ref) > len('refs/heads/'),
            'Source ref must name a nonempty branch under refs/heads/')
    require(re.fullmatch('[0-9a-f]{40}', args.workflow_sha), 'Workflow SHA must be lowercase 40-hex')
    request = {'ref': args.source_ref[len('refs/heads/'):], 'inputs': {
        'owner_compatibility': False,
        'source_sha': source_sha,
        'source_ref': args.source_ref,
        'workflow_sha': args.workflow_sha,
        'runner_label': 'ubuntu-24.04',
        'config_sha': args.transfer_sha256,
        'f283_benchmark': True,
        'f283_approval_bundle': transfer_text}}
    request_bytes = (json.dumps(request, indent=2, allow_nan=False) + '\n').encode('ascii')
    require(len(request_bytes) <= 65535, 'Complete dispatch request exceeds 65535 bytes')
    atomic_bytes(output, request_bytes)
    print('Prepared dispatch request: ' + str(output) + ' (' + str(len(request_bytes))
          + ' bytes, sha256 ' + hashlib.sha256(request_bytes).hexdigest() + ')')
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--self-test', action='store_true')
    parser.add_argument('--proof')
    parser.add_argument('--proof-sha256')
    parser.add_argument('--output')
    parser.add_argument('--case', type=int, choices=[1, 2], help=argparse.SUPPRESS)
    parser.add_argument('--prepare-dispatch', action='store_true')
    parser.add_argument('--transfer')
    parser.add_argument('--transfer-sha256')
    parser.add_argument('--source-ref')
    parser.add_argument('--workflow-sha')
    args = parser.parse_args()
    preparation = [args.prepare_dispatch, args.transfer, args.transfer_sha256, args.source_ref, args.workflow_sha]
    execution = [args.proof, args.proof_sha256, args.case]
    if args.self_test:
        require(not any(preparation + execution + [args.output]), 'Self-test does not accept execution or preparation options')
        print(json.dumps(self_test(), indent=2))
        return 0
    if any(preparation):
        require(not any(execution), 'Dispatch preparation does not accept execution options')
        require(all(preparation),
                'Dispatch preparation requires --transfer, --transfer-sha256, --source-ref, --workflow-sha and --output')
        return prepare_dispatch(args)
    require(all([args.proof, args.proof_sha256, args.output]), 'Explicit reviewed proof/hash and new output directory required')
    return execute(args) if args.case else run_both(args)


if __name__ == '__main__':
    sys.exit(main())
