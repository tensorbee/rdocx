# oxml-chart

ChartML modeling, editing, workbook integration, validation, and backend-neutral chart geometry.

## Use it when

Use this crate when implementing editable OOXML charts or rendering chart
geometry. Use `rpptx` for charts inside a complete presentation.

## Relationship

It uses SpreadsheetML workbooks from `oxml-sml` and feeds backend-neutral
chart geometry into document-family renderers.

## Example

```rust,no_run
use oxml_chart::AxisId;

let category_axis = AxisId::new(10_000_001)?;
let value_axis = AxisId::new(10_000_002)?;
assert_ne!(category_axis, value_axis);
# Ok::<(), oxml_chart::ChartError>(())
```

Add `oxml-chart = "0.11.0"` to your dependencies. See the [chart API](https://docs.rs/oxml-chart) for supported plot families.
