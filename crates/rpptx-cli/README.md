# rpptx-cli

Command-line interface for inspecting, extracting, converting, comparing, replacing, validating, rendering, and outlining PPTX files.

## Use it when

Use the CLI for shell automation. Use `rpptx` when the same operations belong inside a Rust application.

## Relationship

It uses `oxml-cli-support` for shared command conventions and the real `rpptx` facade for document behavior.

## Example

```text
cargo install rpptx-cli --version '^0.11.0'
rpptx inspect deck.pptx --json
rpptx convert deck.pptx --to pdf -o deck.pdf
rpptx thumbnail deck.pptx -o thumbnail.png
```
