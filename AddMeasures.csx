// Tabular Editor C# Script — Add all Ottawa Crime DAX measures
// How to use:
//   1. Open Power BI Desktop with the Ottawa Crime Analysis report
//   2. Launch Tabular Editor from External Tools ribbon
//   3. Go to Advanced Scripting tab (C# Script)
//   4. Paste this entire script and click Run (F5)
//   5. Click Save (Ctrl+S) to push changes back to Power BI Desktop

var table = Model.Tables["Criminal Offences"];

// Helper: add or update a measure
Action<string, string, string> AddMeasure = (name, expression, formatString) =>
{
    Measure m;
    if (table.Measures.Contains(name))
        m = table.Measures[name];
    else
        m = table.AddMeasure(name);
    m.Expression = expression;
    if (!string.IsNullOrEmpty(formatString))
        m.FormatString = formatString;
};

// --- Core Counts ---

AddMeasure("Count of Crimes",
    @"COUNT('Criminal Offences'[Type of Crime])",
    "0");

AddMeasure("Violent Crimes",
    @"CALCULATE (
    COUNT ( 'Criminal Offences'[Type of Crime] ),
    FILTER ( 'Criminal Offences', 'Criminal Offences'[Type of Crime] = ""Violent"" )
)",
    "0");

AddMeasure("Major Crimes",
    @"CALCULATE (
    COUNT ( 'Criminal Offences'[Type of Crime] ),
    FILTER ( 'Criminal Offences', 'Criminal Offences'[Type of Crime] = ""Major"" )
)",
    "0");

// --- Duration & Averages ---

AddMeasure("Date Diff Measure",
    @"DATEDIFF(
    MIN('Criminal Offences'[Occurred Date]),
    MAX('Criminal Offences'[Occurred Date]),
    DAY
)",
    "0");

AddMeasure("Avg Duration",
    @"AVERAGEX (
    'Criminal Offences',
    DATEDIFF (
        MIN ( 'Criminal Offences'[Occurred Date] ),
        MAX ( 'Criminal Offences'[Occurred Date] ),
        DAY
    )
)",
    "0");

AddMeasure("Weeknum measure",
    @"ROUNDDOWN ( [Avg Duration] / 7, 0 )",
    "0");

AddMeasure("Average crime per weekday",
    @"DIVIDE ( [Count of Crimes], [Weeknum measure] )",
    "0.00");

AddMeasure("Avg Date Diff Per Day",
    @"DIVIDE([Count of Crimes], [Avg Duration])",
    "0.00");

// --- Crime Rates ---

AddMeasure("Violent Crime Rate",
    @"DIVIDE([Violent Crimes], [Count of Crimes])",
    "0.00%");

AddMeasure("Major Crime Rate",
    @"DIVIDE ( [Major Crimes], [Count of Crimes] )",
    "0.00%");

// --- Averages per Day ---

AddMeasure("Violent Crimes average per Day",
    @"AVERAGEX(
    KEEPFILTERS(VALUES('Criminal Offences'[Occurred Date].[Day])),
    CALCULATE([Violent Crimes])
)",
    "0.00");

AddMeasure("Major Crimes average per Day",
    @"AVERAGEX(
    KEEPFILTERS(VALUES('Criminal Offences'[Occurred Date].[Day])),
    CALCULATE([Major Crimes])
)",
    "0.00");

// --- Average per Neighbourhood ---

AddMeasure("Crime Rate average per Neighbourhood",
    @"AVERAGEX(
    KEEPFILTERS(VALUES('Criminal Offences'[Neighbourhood Name])),
    CALCULATE([Count of Crimes])
)",
    "0.00");

AddMeasure("Violent Crime Rate average per Neighbourhood",
    @"AVERAGEX(
    KEEPFILTERS(VALUES('Criminal Offences'[Neighbourhood Name])),
    CALCULATE([Violent Crime Rate])
)",
    "0.00%");

// --- YoY% Measures ---

AddMeasure("Count of Crimes YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Count of Crimes],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Count of Crimes] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

AddMeasure("Violent Crimes YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Violent Crimes],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Violent Crimes] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

AddMeasure("Major Crimes YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Major Crimes],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Major Crimes] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

AddMeasure("Violent Crime Rate YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Violent Crime Rate],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Violent Crime Rate] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

AddMeasure("Major Crime Rate YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Major Crime Rate],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Major Crime Rate] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

AddMeasure("Crimes average per Day YoY%",
    @"IF(
    ISFILTERED('Criminal Offences'[Occurred Date]),
    ERROR(""Time intelligence quick measures can only be grouped or filtered by the Power BI-provided date hierarchy or primary date column.""),
    VAR __PREV_YEAR =
        CALCULATE(
            [Average crime per weekday],
            DATEADD('Criminal Offences'[Occurred Date].[Date], -1, YEAR)
        )
    RETURN
        DIVIDE([Average crime per weekday] - __PREV_YEAR, __PREV_YEAR)
)",
    "0.00%");

// --- YoY Diff Display (with arrows) ---

AddMeasure("Diffs Violent Crimes YOY",
    @"SWITCH (
    TRUE (),
    [Violent Crimes YoY%] < 0, UNICHAR ( 9660 ),
    UNICHAR ( 9650 )
)
    & ROUND ( [Violent Crimes YoY%], 4 ) * 100 & "" %""",
    null);

AddMeasure("Diffs Major Crimes YOY",
    @"SWITCH ( TRUE (), [Major Crimes YoY%] < 0, UNICHAR ( 9660 ), UNICHAR ( 9650 ) )
    & ROUND ( [Major Crimes YoY%], 4 ) * 100 & "" %""",
    null);

AddMeasure("Diffs VCR YOY",
    @"SWITCH (
    TRUE (),
    [Violent Crime Rate YoY%] < 0, UNICHAR ( 9660 ),
    UNICHAR ( 9650 )
)
    & ROUND ( [Violent Crime Rate YoY%], 4 ) * 100 & "" %""",
    null);

AddMeasure("Diffs MCR YOY",
    @"SWITCH (
    TRUE (),
    [Major Crime Rate YoY%] < 0, UNICHAR ( 9660 ),
    UNICHAR ( 9650 )
)
    & ROUND ( [Major Crime Rate YoY%], 4 ) * 100 & "" %""",
    null);

// --- Color Measures ---

AddMeasure("Color Measure",
    @"IF (
    'Criminal Offences'[Violent Crime Rate] < .1846
        && 'Criminal Offences'[Major Crime Rate] < .3012,
    ""#008000"",
    IF (
        'Criminal Offences'[Violent Crime Rate] >= .1846
            && 'Criminal Offences'[Major Crime Rate] >= .3012,
        ""#FF0000"",
        IF (
            'Criminal Offences'[Violent Crime Rate] >= .1846
                && 'Criminal Offences'[Major Crime Rate] <= .3012,
            ""#FF7F50"",
            ""#FFFF00""
        )
    )
)",
    null);

AddMeasure("Color Measure for Columns",
    @"IF ( [Violent Crimes] > 0, ""#1E5A8E"", ""#BD890F"" )",
    null);

AddMeasure("Map color Measure",
    @"IF (
    [Count of Crimes YoY%] <= 0,
    ""#008000"",
    IF (
        [Count of Crimes YoY%] > 0
            && [Count of Crimes YoY%] < .1,
        ""#FFFF00"",
        ""#FF0000""
    )
)",
    null);

// --- Slicer / Filter Display Measures ---

AddMeasure("Offence Category Measure",
    @"IF (
    ISFILTERED ( 'Criminal Offences'[Offence Category]),
    CONCATENATEX ( VALUES ( 'Criminal Offences'[Offence Category]), 'Criminal Offences'[Offence Category], "",""),
    ""Select A Type of Crime ""
)",
    null);

AddMeasure("Neighbourhood Name Measure",
    @"IF (
    ISFILTERED ( 'Criminal Offences'[Neighbourhood Name]),
    CONCATENATEX ( VALUES ( 'Criminal Offences'[Neighbourhood Name]), 'Criminal Offences'[Neighbourhood Name], "",""),
    ""Select A Neighbourhood ""
)",
    null);

AddMeasure("Intersection Level Measure",
    @"IF (
    ISFILTERED ( 'Criminal Offences'[Intersection] ),
    CONCATENATEX ( VALUES ( 'Criminal Offences'[Intersection]), 'Criminal Offences'[Intersection], "",""),
    ""Select An Intersection ""
)",
    null);

// --- Transparency / Conditional Display Measures ---

AddMeasure("Make Transparent Major Crime Rate",
    @"If(IsBlank([Major Crime Rate]), ""VC Selected"", [Major Crime Rate])",
    null);

AddMeasure("Make Transparent Major Crimes YOY%",
    @"If(IsBlank([Major Crimes YoY%]), ""VC Selected"", [Diffs Major Crimes YOY])",
    null);

AddMeasure("Make Transparent Violent Crime Rate",
    @"If(IsBlank([Violent Crime Rate]), ""MC Selected"", FORMAT([Violent Crime Rate], ""0.00%""))",
    null);

AddMeasure("Make Transparent Violent Crimes YOY%",
    @"If(IsBlank([Violent Crimes YoY%]), ""MC Selected"", [Diffs Violent Crimes YOY])",
    null);

// Done
Info("All 34 measures added to 'Criminal Offences' table. Press Ctrl+S to save back to Power BI Desktop.");
