# AstroSage Call Center Performance Analysis

## Project Overview

This project analyzes AstroSage consultation activity using Excel to
understand call and chat volumes, revenue, astrologer workload, customer
ratings, repeat usage, and operational performance. The workbook
contains the source data, a cleaned dataset, objective-question analysis
sheets, supporting pivot tables/charts, and a dashboard. A PowerPoint
presentation summarizes selected findings and recommendations.

The analysis covers consultation records from **December 2023 to January
2024**. The working dataset contains **28,027 records**, with 35
original attributes and 7 added helper columns in the cleaned sheet.

## Project Goals

-   Clean and prepare the raw consultation dataset for analysis.
-   Track daily and monthly consultation volume.
-   Compare revenue by consultation type and month.
-   Understand call-centre workload across astrologers.
-   Measure repeat-user activity.
-   Explore the relationship between consultation duration and customer
    ratings.
-   Compare customer satisfaction across months and astrologers.
-   Analyze platform/website distribution, consultation status, hourly
    patterns, and rating distributions.
-   Turn the analysis into practical recommendations for service
    quality, staffing, and revenue resilience.

## Files Included

  ---------------------------------------------------------------------------
  File                                    Description
  --------------------------------------- -----------------------------------
  `Untitled spreadsheet.xlsx`             Main Excel workbook containing raw
                                          data, cleaned data,
                                          objective-question analysis sheets,
                                          pivot tables, charts, and a
                                          dashboard.

  `Doc Template (Astrosage) (1) 1.docx`   Written project solution with
                                          question statements, final answers,
                                          analysis approaches, and reported
                                          results.

  `AstroSage_Analytics_Stylish.pptx`      Presentation summarizing the
                                          business problem, selected charts,
                                          key findings, and recommendations.
  ---------------------------------------------------------------------------

## Dataset Summary

  -----------------------------------------------------------------------
  Item                                Description
  ----------------------------------- -----------------------------------
  Records                             28,027 consultation records

  Original attributes                 35 columns

  Cleaned attributes                  42 columns, including 7 helper
                                      columns

  Period                              December 2023--January 2024

  Unique users                        10,344, as reported in the analysis
                                      notes

  Astrologers/gurus                   The notes report 151 unique gurus
                                      in one analysis and 149 agents in
                                      the workload calculation; reconcile
                                      the counting rule before final
                                      reporting

  Main channels                       Call and Chat, with a small number
                                      of other/free consultation types

  Main output                         Excel analysis and dashboard,
                                      written report, and presentation
  -----------------------------------------------------------------------

The dataset includes consultation identifiers, user and guru fields,
consultation type, website/platform, status, timing/duration fields,
ratings, and financial fields. Exact column availability should be
checked in the workbook before reusing formulas.

## Workbook Structure

The workbook contains these sheets:

-   **`data`** --- original source records.
-   **`Cleaned data`** --- prepared dataset with helper columns.
-   **`Q-4 I Objective`** through **`Q-13 I Objective`** --- analysis
    for the objective questions.
-   **`Q-1 I Subjective`**, **`Q-3 I Subjective (Optional)`**, and
    additional subjective-question sheets --- business interpretation
    and recommendations.
-   **`website distrubition`** --- platform/website distribution
    analysis.
-   **`Dashboard`** --- consolidated visual reporting.

Sheet names are retained as supplied and may include spelling or naming
inconsistencies.

## Data Cleaning and Preparation

The project notes describe the following preparation steps:

1.  Copied the original data into a separate **Cleaned data** worksheet
    to preserve the source.
2.  Added date helper fields:
    -   `Date`
    -   `Month`
    -   `Year`
    -   `Day`
3.  Created `Clean_Rating` to treat rating `0` as blank for analyses of
    submitted ratings. The notes associate zero ratings with failed or
    unrated sessions.
4.  Created `Duration_Min` by converting seconds to minutes based on
    consultation type, using chat duration for Chat and call duration
    for Call.
5.  Added a `Repeat_User` helper field to label users with multiple
    consultation records as `Repeat`, and users with one record as
    `New`.
6.  Checked for duplicate rows; the notes report no duplicate rows.
7.  Reviewed missing values and treated blank status fields as
    potentially valid depending on consultation type.

Example formulas recorded in the project notes:

``` excel
=INT(P2)
=TEXT(date_cell,"mmm")
=IF(rating_cell=0,"",rating_cell)
=IF(consultation_type="Chat",chat_seconds/60,IF(consultation_type="Call",call_duration/60,""))
=IF(COUNTIF($B$2:$B$28028,B2)>1,"Repeat","New")
```

These examples use illustrative cell references. Confirm the actual
column positions in the current workbook before using them. In
particular, verify the source date, rating, chat-duration,
call-duration, and consultation-type columns.

## Analysis Questions and Topics

The workbook/report covers these principal questions and analyses:

1.  Number of tables/data sheets and total attributes.
2.  Dataset cleaning and handling of inconsistent or missing values.
3.  Day-by-day consultation volume and average daily volume.
4.  Highest- and lowest-volume months.
5.  Monthly financial totals, net revenue, and astrologer payouts.
6.  Average consultations handled per agent per day.
7.  Repeat users and the share of consultations from repeat users.
8.  Revenue by consultation type.
9.  Consultation counts by user and guru.
10. Relationship between consultation duration and customer
    satisfaction.
11. Highest- and lowest-rated gurus.
12. Average customer satisfaction by month.
13. Identification of categorical, continuous/numerical, and ID/text
    attributes.
14. Website/platform distribution.
15. Consultation status distribution.
16. Call distribution by hour and daily activity trends.
17. Rating-wise user/guru distributions and recommendations for
    improving performance.

The dashboard and presentation visualize selected topics using pivot
tables, charts, and summary metrics.

## Reported Key Findings

The following figures are reported in the supplied analysis notes. They
should be validated against the current workbook before submission or
external use.

### Consultation volume

-   Total consultations: **28,027**.
-   December 2023: **26,488 consultations**.
-   January 2024: **1,539 consultations**, described in the notes as
    only three days of data.
-   Average daily consultation volume: **824 sessions** across 34 days.
-   The reported peak day is **15 December 2023**, with **1,324
    sessions**.
-   The reported low day is **1 January 2024**, with **325 sessions**.

### Revenue and payouts

  Consultation type/month         Reported value
  ----------------------------- ----------------
  Call revenue                         ₹1,68,442
  Chat revenue                           ₹45,495
  Public live call revenue                   ₹51
  Complimentary consultations                 ₹0
  **Total net revenue**            **₹2,13,987**
  Total amount charged                 ₹2,14,066
  Astrologer payouts                     ₹99,147

Monthly financial figures in the notes:

  Month             Total amount   Net revenue   Astrologer payouts
  --------------- -------------- ------------- --------------------
  December 2023        ₹2,02,239     ₹2,02,215              ₹93,786
  January 2024           ₹11,827       ₹11,773               ₹5,360

### Workload and repeat usage

-   Average workload is reported as **5.53 consultations per agent per
    day**, using 28,027 consultations, 149 agents, and 34 days.
-   Total unique users are reported as **10,344**.
-   The notes report **22,745 repeat consultations** and **5,282
    new-user consultations**.
-   Consultations from repeat users are reported as **81.15% of total
    calls**; the share of users classified as repeat is reported as
    **48.94%**.
-   Astro Krishaa is reported as the top guru by consultation count,
    with **1,580 consultations**.

### Customer satisfaction

-   Overall average rating is reported as **3.96 on an 8-point scale**
    when zero ratings are excluded.
-   Reported duration-rating correlation: **0.056**, suggesting a very
    weak linear relationship.
-   The 2--5 minute duration group is reported to have an average rating
    of **4.19/8**.
-   Highest-rated guru in one result: Astro Saraswat, **6.31/8**, based
    on 18 ratings.
-   A more sample-supported high performer is reported as Daljit Kaur,
    **6.06/8**, based on 111 ratings.
-   Lowest-rated guru in one result: Acharya Divyansh, **1.87/8**, based
    on 23 ratings.
-   Monthly average ratings excluding zero ratings are reported as
    **3.93/8 for December 2023** and **4.67/8 for January 2024**.

## Tools and Excel Techniques

-   Microsoft Excel
-   Data cleaning and helper columns
-   Pivot tables and pivot charts
-   Date grouping and month/day analysis
-   `COUNTA`, `COUNTIF`, `AVERAGE`, `IF`, `IFERROR`/`IFNA` where needed
-   `CORREL` for duration-rating correlation
-   Revenue aggregation using `SUM`
-   Sorting, filtering, and ranking
-   Dashboard design and visual storytelling

## How to Use the Workbook

1.  Open `Untitled spreadsheet.xlsx` in Microsoft Excel.
2.  Review the `data` sheet to understand the source fields.
3.  Review `Cleaned data` and verify the helper columns and formulas.
4.  Open each `Q-...` sheet to inspect the pivot tables, calculations,
    and charts for the relevant question.
5.  Open `Dashboard` to review the consolidated visual analysis.
6.  Refresh pivot tables if the underlying data changes.
7.  Check all date, status, rating, and duration definitions before
    interpreting results.
8.  Compare workbook results with the written report and PowerPoint.
9.  Validate the figures in the report and slides before submitting or
    sharing the project.

## Recommendations

The project presentation recommends several improvement areas:

-   **Improve astrologer quality:** provide targeted coaching, review
    low-rating patterns, and share best practices from consistently
    strong performers.
-   **Improve completion rate:** investigate failed, busy, cancelled, or
    otherwise incomplete sessions and address the causes by consultation
    type.
-   **Plan staffing around demand:** use hourly, daily, and weekday
    trends to schedule sufficient astrologer availability during peaks.
-   **Increase feedback coverage:** make rating prompts easy to complete
    after consultations and monitor the unrated-session share.
-   **Reduce platform concentration risk:** understand why activity is
    concentrated on one platform and evaluate opportunities to improve
    the experience on other channels.
-   **Monitor revenue mix:** calls are reported as the largest revenue
    source, so improve call quality while maintaining healthy
    performance across chat and other services.
-   **Build a recurring KPI review:** track volume, completion, revenue,
    payouts, ratings, repeat usage, and workload over time.

These are proposed actions based on the analysis, not guarantees of
future outcomes.

## Validation Notes and Limitations

Before treating the findings as final, review these points:

-   **Agent count differs:** the workbook notes mention 151 unique gurus
    in one place and use 149 agents in the average-workload calculation.
    Decide whether the counts differ because of filtering, naming, or
    distinct-count methodology.
-   **Rating scale:** the Excel analysis uses an 8-point scale, while
    the PowerPoint contains a slide describing average rating as
    `3.8/5`. Standardize the scale and labels before presenting.
-   **Completion rate:** the PowerPoint reports 19.7% completion (5,535
    of 28,027 sessions). Confirm which status values count as completed
    and whether the denominator includes every consultation type.
-   **Date/chart anomalies:** one slide labels a low-activity date as
    "May 1901," which appears inconsistent with the stated December
    2023--January 2024 analysis period. Check the date field and chart
    source before using this claim.
-   **Guru names and rankings:** names and counts differ slightly across
    project notes and presentation slides. Refresh the pivots and use
    one consistent source of truth.
-   **Repeat-user calculations:** distinguish between repeat users
    (distinct people with multiple consultations) and repeat
    consultations (individual records from returning users). These are
    different measures.
-   **Zero ratings:** excluding zero ratings is appropriate for an
    average of submitted ratings only if zero means "unrated/failed,"
    not a valid customer score. Report the zero/unrated count
    separately.
-   **Duration correlation:** correlation does not establish that call
    duration causes a higher or lower rating. Failed sessions,
    consultation type, guru, and other factors may affect both.
-   **Financial definitions:** confirm the meanings of `amount`,
    `netAmount`, and `astrologersEarnings` before describing them as
    gross revenue, net revenue, or cost.
-   **January comparison:** January contains only a partial month in the
    supplied notes, so it should not be compared with a full December
    without a caveat.
-   **Presentation consistency:** align the slide deck with the workbook
    and written report, including rating scale, completion definitions,
    date ranges, and guru names.

## Future Improvements

-   Add a data dictionary defining every important field.
-   Automate cleaning and validation with Power Query or a repeatable
    import process.
-   Create a consistent KPI definitions sheet.
-   Add slicers for date, website/platform, and consultation type where
    supported by the dashboard design.
-   Add clear labels for failed, busy, completed, and unrated sessions.
-   Compare average ratings with sample sizes so very small groups are
    not over-interpreted.
-   Add a month-to-date or per-day comparison to account for partial
    months.
-   Reconcile all counts and names across the workbook, report, and
    slides.
-   Export a final, validated dashboard and presentation after
    refreshing the pivot tables.

## Conclusion

This project uses Excel-based data preparation, pivot-table analysis,
charts, and dashboard reporting to examine AstroSage's consultation
operations. It highlights consultation volume patterns, revenue by
service type, astrologer workload, repeat usage, and customer
satisfaction. Once the noted differences and chart anomalies are
reconciled, the workbook can support more reliable operational decisions
around quality improvement, scheduling, customer feedback, and revenue
management.

## Author

**Sushant Kumar**

------------------------------------------------------------------------

*This README documents the supplied project files and the findings
stated in the accompanying notes. Reported metrics have not been
independently recalculated from every source row as part of preparing
this README.*
