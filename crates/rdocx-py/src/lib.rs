//! Python bindings for the rdocx facade.

mod document;
mod formatting;
mod paragraph;
mod run;
mod table;

use pyo3::exceptions::{PyIndexError, PyRuntimeError};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyType};

use oxml_py_support::StaleElementError;

use document::PyDocument;
use formatting::{PyFont, PyParagraphFormat};
use paragraph::{PyParagraph, PyParagraphCollection};
use run::{PyRun, PyRunCollection};
use table::{
    PyCell, PyCellCollection, PyCellParagraphCollection, PyRow, PyRowCollection, PyTable,
    PyTableCollection,
};

pub(crate) fn normalize_index(index: isize, len: usize, kind: &str) -> PyResult<usize> {
    let normalized = if index < 0 {
        len as isize + index
    } else {
        index
    };
    if normalized < 0 || normalized >= len as isize {
        return Err(PyIndexError::new_err(format!("{kind} index out of range")));
    }
    Ok(normalized as usize)
}

pub(crate) fn length_object(py: Python<'_>, value: rdocx::Length) -> PyResult<Py<PyAny>> {
    py.import("rdocx")?
        .getattr("Length")?
        .call1((value.to_emu(),))
        .map(Bound::unbind)
}

pub(crate) fn enum_object(py: Python<'_>, name: &str, value: i32) -> PyResult<Py<PyAny>> {
    py.import("rdocx")?
        .getattr(name)?
        .call1((value,))
        .map(Bound::unbind)
}

fn public_error(py: Python<'_>, class_name: &str, message: String) -> PyErr {
    let exception_type = py
        .import("rdocx")
        .and_then(|module| module.getattr(class_name))
        .and_then(|class| class.cast_into::<PyType>().map_err(Into::into));

    match exception_type {
        Ok(class) => PyErr::from_type(class, (message,)),
        Err(_) => PyRuntimeError::new_err(message),
    }
}

pub(crate) fn stale_to_pyerr(py: Python<'_>, error: StaleElementError) -> PyErr {
    public_error(py, "StaleElementError", error.to_string())
}

pub(crate) fn rdocx_to_pyerr(py: Python<'_>, error: rdocx::Error) -> PyErr {
    let class_name = match &error {
        rdocx::Error::Opc(_)
        | rdocx::Error::Io(_)
        | rdocx::Error::NoDocumentPart
        | rdocx::Error::UnavailableImageDimensions { .. } => "PackageError",
        rdocx::Error::Oxml(_) => "XmlError",
        rdocx::Error::Layout(_) | rdocx::Error::Pdf(_) | rdocx::Error::Raster(_) => "LayoutError",
        rdocx::Error::Rtf { .. }
        | rdocx::Error::Html { .. }
        | rdocx::Error::Odt { .. }
        | rdocx::Error::Other(_) => "RdocxError",
    };
    public_error(py, class_name, error.to_string())
}

#[pymodule]
fn _rdocx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyDocument>()?;
    module.add_class::<PyParagraph>()?;
    module.add_class::<PyParagraphCollection>()?;
    module.add_class::<PyRun>()?;
    module.add_class::<PyRunCollection>()?;
    module.add_class::<PyFont>()?;
    module.add_class::<PyParagraphFormat>()?;
    module.add_class::<PyTable>()?;
    module.add_class::<PyTableCollection>()?;
    module.add_class::<PyRow>()?;
    module.add_class::<PyRowCollection>()?;
    module.add_class::<PyCell>()?;
    module.add_class::<PyCellCollection>()?;
    module.add_class::<PyCellParagraphCollection>()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pyo3::ffi::c_str;

    #[test]
    fn layout_error_maps_to_the_exact_public_layout_error_class() {
        Python::initialize();
        Python::attach(|py| -> PyResult<()> {
            let package = PyModule::from_code(
                py,
                c_str!(
                    "class RdocxError(Exception):\n    pass\n\nclass LayoutError(RdocxError):\n    pass\n"
                ),
                c_str!("rdocx_test.py"),
                c_str!("rdocx"),
            )?;
            py.import("sys")?
                .getattr("modules")?
                .set_item("rdocx", &package)?;
            let expected = package.getattr("LayoutError")?.cast_into::<PyType>()?;

            let error = rdocx::Error::Layout(oxml_layout::LayoutError::Layout(
                "classifier regression".to_string(),
            ));
            let raised = rdocx_to_pyerr(py, error);

            assert!(raised.get_type(py).is(&expected));
            Ok(())
        })
        .unwrap();
    }

    #[test]
    fn import_errors_map_to_the_generic_public_error_class() {
        Python::initialize();
        Python::attach(|py| -> PyResult<()> {
            let package = PyModule::from_code(
                py,
                c_str!("class RdocxError(Exception):\n    pass\n"),
                c_str!("rdocx_test.py"),
                c_str!("rdocx"),
            )?;
            py.import("sys")?
                .getattr("modules")?
                .set_item("rdocx", &package)?;
            let expected = package.getattr("RdocxError")?.cast_into::<PyType>()?;

            for error in [
                rdocx::Error::Html {
                    location: "body[0]".to_string(),
                    message: "invalid HTML".to_string(),
                },
                rdocx::Error::Odt {
                    part: Some("content.xml".to_string()),
                    offset: 0,
                    message: "invalid ODT".to_string(),
                },
            ] {
                assert!(rdocx_to_pyerr(py, error).get_type(py).is(&expected));
            }
            Ok(())
        })
        .unwrap();
    }
}
