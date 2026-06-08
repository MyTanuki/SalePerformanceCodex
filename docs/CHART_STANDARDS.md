# Chart Standards

This document defines reusable chart display standards for the Sale Performance Dashboard.

## Target vs Actual Bar

Use this pattern whenever a chart compares a target value with an actual, won, or achieved value in the same period.

### Meaning

- `Target` is the reference frame.
- `Actual`, `Won`, or achieved value is the measured result.
- Both values must be drawn on the same scale.
- If Actual is greater than Target, the Actual bar must be longer than the Target frame.
- If Actual is less than Target, the Actual bar must be shorter than the Target frame.

### Visual Style

- Draw `Target` as a transparent outline frame.
- Draw `Actual` / `Won` as a solid filled bar.
- The Target frame should be taller than the Actual bar.
- The Actual bar should be vertically centered inside the Target frame.
- The Actual bar should have a straight right edge, not a rounded end.
- The Target outline must remain visible even when Actual is equal to or greater than Target.

### Layering

Use three layers:

1. `Target base`: bottom layer, same width as Target, used as the target hover area.
2. `Actual bar`: middle layer, solid fill, used as the actual hover area.
3. `Target outline`: top layer, visible outline only, with pointer events disabled so it does not block Actual hover.

### Tooltip Rules

- Hovering over the filled Actual bar must show the Actual / Won value.
- Hovering over the Target frame outside the filled bar must show the Target value.
- The Target outline overlay must not intercept pointer events.

### Color Rules

- Total Sales charts use the Total Sales blue theme.
- Renew and recurring revenue charts use the Renew green theme.
- New Sales charts use the New Sales purple theme.
- Do not mix multiple semantic colors in the same Target vs Actual bar unless the chart compares multiple categories.

### Gap Rule

When showing a gap beside the chart:

- `Gap = Actual - Target`
- Actual greater than Target must show a positive gap.
- Actual less than Target must show a negative gap.
- Positive and negative signs should be displayed explicitly.

## Current Uses

This standard applies to:

- `Overview > Target vs Actual by Month`
- `Revenue Analysis > Revenue Performance > Monthly Performance Chart`
- `Revenue Analysis > Revenue Analysis > Monthly Revenue Schedule`

