# Filtering Rules

Filtering rules are a common source of regressions. Before changing filters, identify which business question the view answers.

## Global View Filter

The global view filter includes:

- view mode
- sale group
- sale owner
- keyword search
- New/Renew/Total type buttons

It generally affects Sales Performance, Overview, All Deals, Pipeline Risk, Revenue Analysis, Revenue Performance, Revenue Distribution, and Invoice Analysis.

## Global Period Filter

The global period filter supports:

- Year
- Quarter
- Month
- Range

Week filtering was removed from the global period filter. ISO Week belongs to New Deals.

## Date Fields By Area

Sales Performance:

- Uses Expected Close Date for period tracking.

All Deals:

- Shows deal details in the selected sales period.
- Must respect global filter and its own AND/NOT table search.

Pipeline Risk:

- Uses open counted deals and sales tracking dates.

New Deals:

- Uses Created Date.
- Uses ISO Week.
- Has its own New Deals search.

Revenue Analysis:

- Uses revenue schedule entries generated from contract/service rules.
- Date filtering applies to revenue entry month, not raw Expected Close Date.

Invoice Analysis:

- Uses Posting Date mode or Service Period mode depending on the selected invoice mode.

## Search Behavior

Keyword search:

- Should filter visible deal/revenue/invoice datasets where the business question is about a selected customer/deal.
- Should not distort Recurring Target.

Detail AND/NOT search:

- Applies to detail tables where provided.
- Should not change Recurring Target.

New Deals search:

- Lives inside New Deals.
- Should not depend on global keyword search unless intentionally changed.

## Recurring Target Search Boundary

Recurring Target should be controlled by:

- selected year
- sale group
- sale owner
- New/Renew type

It should not be controlled by:

- global keyword search
- Revenue Detail AND/NOT search
- detail table column filters

Reason: Recurring Target is a planning baseline, not a filtered search result.

