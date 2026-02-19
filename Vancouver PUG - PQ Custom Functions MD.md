
# Custom Functions in Power Query (M Language)
> Build reusable, maintainable, and powerful data transformations with Power Query in Power BI 
> All levels welcome






# About Me


<img width="455" height="592" alt="image" src="https://github.com/user-attachments/assets/89f64e7d-405d-471c-bbeb-9de866b68ad2" />
<br /> 
<br /> 


- **Project Manager** (PMP, PMI-ACP) 
- **Microsoft Certified Professional** (DP-600, PL-300, MCT, MCSA, MCSE, MOS Master)
- **Romania PBI and Modern Excel UG Founder** — [https://www.meetup.com/romaniapug](https://www.meetup.com/romaniapug)
- **Microsoft MVP** M365 (**Excel**) & Data Platform (**Power BI**)
- **Proud dad and husband

---

🔗 [linkedin.com/in/cristian-angyal](https://linkedin.com/in/cristian-angyal)  
🔗 [https://linktr.ee/cristianangyal](https://linktr.ee/cristianangyal)






<br /> 
<br /> 
<br /> 
<br /> 


---

## Agenda 

| #   | Section                      |
| --- | ---------------------------- |
| 01  | Why Custom Functions?        |
| 02  | Syntax & Structure           |
| 03  | Parameters & Types           |
| 04  | Real-World Use Cases (Demos) |
| 05  | Error Handling               |
| 06  | Performance Tips             |
| 07  | Reusing Your Functions       |
| 08  | Q&A & Wrap-Up                |


<br /> 
<br /> 
<br /> 
<br /> 


<br /> 
<br /> 
<br /> 
<br /> 


<br /> 
<br /> 
<br /> 
<br /> 

        
          
---

## 01 · Why Custom Functions?
*From copy-paste chaos to reusable, elegant M code*

---

### The Problem: Copy-Paste Madness

❌ **Duplicate logic** — Same transformation repeated in 20 queries. One bug → 20 fixes.

❌ **Hard to maintain** — Next month you won't remember what that nested `List.Transform` was doing.

❌ **Not testable** — No easy way to unit-test an inline transformation step.

✅ **Custom Functions solve all of this** — Write once, call everywhere. Self-documenting. Parameterized.

---

## 02 · Syntax & Structure
*The anatomy of an M function*

---

### Function Anatomy

**The M Function Pattern:**

```m
(parameter1 as type, parameter2 as type)
    as returnType =>
let
    step1 = ...,
    step2 = ...
in
    step2
```

① Parameters list — just like any programming language  
② Optional return type — recommended for documentation  
③ `let…in` block — exactly the same as a regular query!  
④ The final expression after `in` is the return value

---

**Simplest Examples:**

```m
// No parameters
() => "Hello, Power Query!"

// One parameter
(name as text) as text =>
    "Hello, " & name & "!"

// Arithmetic
(x as number, y as number) as number =>
    x + y
```

---

### 3 Ways to Create a Custom Function

**Method 1 — Blank Query → Paste M code**
Most flexible. Full control over code. Best for complex functions.
```m
// Right-click in Query pane
// → New Query → Blank Query
// → Open Advanced Editor
// Paste your function code
```

**Method 2 — Convert Query to Function**
Right-click any existing query → *Create Function*. Power Query wraps it automatically.
```m
// Right-click a query step
// "Create Function" option
// PQ wraps your query body
// in a function signature
```

**Method 3 — Define inline in another query**
Define and immediately invoke — great for one-off uses within a complex query.
```m
let
    MyFn   = (x as number) => x * 2,
    Result = MyFn(5)  // = 10
in
    Result
```

---

## 03 · Parameters & Types
*Make your functions safe and self-documenting*

---

### M Type System — What You Need to Know

**Primitive Types:**

| Type | Example / Notes |
|------|----------------|
| `text` | `"Hello"` |
| `number` | `42`, `3.14` |
| `logical` | `true`, `false` |
| `date` | `#date(2025,1,1)` |
| `datetime` | `#datetime(...)` |
| `duration` | `#duration(1,0,0,0)` |
| `list` | `{1, 2, 3}` |
| `record` | `[Name="Ana"]` |
| `table` | any table |
| `function` | a function value |
| `nullable X` | X or null allowed |

**Typed vs. Untyped Parameters:**

```m
// Untyped — anything works
(x) => x * 2

// Typed — clear contract
(x as number) as number =>
    x * 2

// Nullable — allows null
(x as nullable text) =>
    if x = null then ""
    else Text.Upper(x)

// Optional — can be omitted
(name as text, prefix as optional text) =>
    let p = prefix ?? "Mr."
    in p & " " & name
```

---

### Self-Documenting Functions with Metadata

Power Query can show rich documentation for your functions — just add metadata.

```m
let
    // Define the function body
    fnBody = (startDate as date, endDate as date,
              label as optional text) as table =>
    let
        days   = Duration.Days(endDate - startDate),
        lbl    = label ?? "Range",
        result = #table({"Label","Days"}, {{lbl, days}})
    in  result,

    // Attach documentation metadata
    fnType = type function
        (startDate as (type date meta [
            Documentation.FieldCaption      = "Start Date",
            Documentation.FieldDescription  = "The first date of the range"
         ]),
         endDate as (type date meta [
            Documentation.FieldCaption = "End Date"
         ]),
         optional label as (type text meta [
            Documentation.FieldCaption = "Label",
            Documentation.SampleValues  = {"FY2025","Q1"}
         ])) as table,

    documented = Value.ReplaceType(fnBody, fnType)
in
    documented
```

---

## 04 · Real-World Use Cases
*Live demos — from simple to sophisticated*

---

### Demo 1 — Text Cleaning & Normalization

**Problem:** Product codes arrive inconsistently — spaces, dashes, wrong case. Clean them in one reusable function.

```m
// Query Name: fn_CleanProductCode
(rawCode as nullable text) as text =>
let
    safe   = rawCode ?? "",
    noSpc  = Text.Remove(safe, {" ", Character.FromNumber(9)}),
    clean  = Text.Remove(noSpc, {"-", "."}),
    result = Text.Upper(Text.Trim(clean))
in
    result

// Usage in any table:
// = Table.TransformColumns(Source, {{"ProductCode", fn_CleanProductCode}})
```

**Before → After:**

| Raw Input | Result |
|-----------|--------|
| ` abc-123 ` | `ABC123` |
| `XY Z.456` | `XYZ456` |
| `null` | *(empty string)* |
| `PQ-2024.Q1` | `PQ2024Q1` |
| `hello world` | `HELLOWORLD` |

---

### Demo 2 — API Pagination with a Function

**Problem:** An API returns max 100 rows per page. Call it repeatedly and combine all pages.

```m
// fn_GetPage — fetches a single page
(baseUrl as text, pageNum as number) as table =>
let
    url  = baseUrl & "?page=" & Number.ToText(pageNum),
    raw  = Web.Contents(url, [Headers=[
               Authorization = "Bearer " & ApiToken]]),
    json = Json.Document(raw),
    tbl  = Table.FromList(json[data], Splitter.SplitByNothing()),
    exp  = Table.ExpandRecordColumn(tbl, "Column1",
               Record.FieldNames(tbl{0}[Column1]))
in  exp

// Orchestrator query — calls fn_GetPage for pages 1..N
let
    totalPages = 5,
    pages      = {1..totalPages},
    tables     = List.Transform(pages, each fn_GetPage(BaseUrl, _)),
    combined   = Table.Combine(tables)
in  combined
```

> 💡 `List.Transform` is your **for-loop** in Power Query.

---

### Demo 3 — Fiscal Year / Period Labeling

**Business context:** Fiscal quarters don't always align with calendar Q1–Q4. Encode the rule once.

```m
// fn_FiscalPeriod
(dt as nullable date, fiscalYearStart as optional number) as record =>
let
    safe     = dt ?? Date.From(DateTime.LocalNow()),
    startMo  = fiscalYearStart ?? 1,       // default: January
    mo       = Date.Month(safe),
    yr       = Date.Year(safe),
    shiftMo  = Number.Mod(mo - startMo + 12, 12) + 1,
    quarter  = Number.RoundUp(shiftMo / 3),
    fiscalYr = if mo >= startMo then yr else yr - 1,
    label    = "Q" & Text.From(quarter) & "-FY" & Text.From(fiscalYr)
in
    [ FiscalYear = fiscalYr, FiscalQ = quarter, Label = label ]

// Returns a record → expand into multiple columns at once:
// = Table.ExpandRecordColumn(
//     Table.AddColumn(src, "FP", each fn_FiscalPeriod([Date], 4)),
//     "FP", {"FiscalYear","FiscalQ","Label"})
```

> 💡 Returning a **record** instead of a scalar lets you produce multiple columns from a single function call.

---

## 05 · Error Handling
*Write functions that fail gracefully*

---

### Error Handling — 4 Essential Patterns

**Pattern 1 — `try...otherwise`**
Returns a fallback value if any error occurs.
```m
(x as text) =>
    try Number.FromText(x)
    otherwise 0
```

**Pattern 2 — Null guard (`??`)**
Null coalescing — handle null inputs safely.
```m
(val as nullable text) =>
    let safe = val ?? ""
    in Text.Upper(safe)
```

**Pattern 3 — `try` with error details**
Inspect the error record for logging or auditing.
```m
(x as text) =>
    let r = try Number.FromText(x)
    in if r[HasError]
       then r[Error][Message]
       else Text.From(r[Value])
```

**Pattern 4 — `error` keyword**
Throw a meaningful custom error for invalid inputs.
```m
(x as number) as number =>
    if x < 0
    then error Error.Record("InvalidInput", "Must be >= 0", x)
    else x
```

> ⚠️ **A function without null handling WILL break in production.**

---

## 06 · Performance Tips
*Write functions that scale — not just work*

---

### Performance — What to Know

⚡ **Avoid calling functions inside Add Column on large tables**
Each row triggers a new M evaluation. For > 100K rows, consider native column operations (`Text.Upper`, etc.) that fold to the source.

🔁 **Query Folding — your best friend**
If the function uses connectors that support folding (SQL Server, SharePoint, etc.), keep transformations foldable. Avoid custom M steps that break the fold.

📦 **Cache expensive lookups**
Don't call `Web.Contents()` inside a row-level function! Fetch once, store in a variable, then pass the result into your function.

🧪 **Test with small data first**
Use `Table.FirstN(source, 100)` while developing. Power Query evaluates lazily — your function may be called more times than you expect.

---

## 07 · Reusing Your Custom Functions
*From this workbook to every workbook*

---

### Method 1 — Copy & Paste the Query

The simplest way. Works in both Power BI and Excel.

**Steps:**
1. In Power Query Editor, right-click the function query in the Queries pane
2. Select **Copy**
3. Open the destination workbook/report → Power Query Editor
4. Right-click in the Queries pane → **Paste**

✅ Instant. No setup required.  
⚠️ Each copy is independent — changes in the source don't propagate.

---

### Method 2 — Copy the M Code

Share the raw M code as text — via email, a wiki, a Teams message, or a GitHub Gist.

**To extract the M code:**
1. Open the function query in Power Query Editor
2. Home → **Advanced Editor**
3. Select All → Copy

**To import it:**
1. Blank Query → Advanced Editor → Paste → Done

```m
// Example: share this snippet with a colleague
(rawCode as nullable text) as text =>
let
    safe   = rawCode ?? "",
    clean  = Text.Remove(Text.Upper(Text.Trim(safe)), {"-","."," "})
in
    clean
```

✅ Easy to share across teams and platforms.  
✅ Version-controllable in Git.  
⚠️ Still manual — recipient must paste and maintain separately.

---

### Method 3 — External Function Library (The Pro Approach)

Store all your functions in a **single source of truth** — one Excel file or a SharePoint/OneDrive location — and reference them from any workbook or report.

**Architecture:**

```
📁 FunctionLibrary.xlsx  (or a blank Power BI file)
    └── fn_CleanProductCode
    └── fn_FiscalPeriod
    └── fn_GetPage
    └── fn_ConvertCurrency
```

**How to connect from any report:**

```m
// In your report's Power Query — reference the library file
let
    Source   = Excel.Workbook(
                   File.Contents("\\shared\BI\FunctionLibrary.xlsx"),
                   null, true),
    fnTable  = Source{[Item="fn_CleanProductCode", Kind="Sheet"]}[Data],
    // Or load as a function directly if stored as a query:
    MyFn     = Expression.Evaluate(
                   fnTable{0}[MCode],
                   #shared)
in
    MyFn
```

**Even simpler — SharePoint as the library host:**

```m
// Load function M code stored as text in a SharePoint list
let
    Source  = SharePoint.Tables("https://yourorg.sharepoint.com/sites/BI",
                  [ApiVersion = 15]),
    fnLib   = Source{[Title="FunctionLibrary"]}[Items],
    fnRow   = Table.SelectRows(fnLib, each [FunctionName] = "fn_CleanProductCode"){0},
    MyFn    = Expression.Evaluate(fnRow[MCode], #shared)
in
    MyFn
```

✅ **Single source of truth** — update once, all reports benefit.  
✅ Works across the entire team.  
✅ Can be version-controlled and audited.  
⚠️ Requires `Expression.Evaluate` — needs **trust settings** enabled in Power BI (`File → Options → Security → Allow...`).

---

### Comparison

| | Copy & Paste | Share M Code | Function Library |
|--|:--:|:--:|:--:|
| Setup effort | 🟢 None | 🟢 Minimal | 🟡 Some |
| Stays in sync | ❌ No | ❌ No | ✅ Yes |
| Team sharing | 🟡 Manual | 🟡 Manual | 🟢 Automatic |
| Version control | ❌ No | 🟢 Via Git | 🟢 Via Git/SharePoint |
| Best for | Quick reuse | Sharing snippets | Enterprise / teams |

---

## Key Takeaways

1. Functions use the same `let…in` syntax as queries — nothing new to learn.
2. Always type your parameters and handle nulls — safety first.
3. Return records to get multiple columns from one function call.
4. `Table.TransformColumns` and `List.Transform` are your power tools.
5. Test with `Table.FirstN`; watch for query folding to stay performant.
6. Add metadata to make functions feel like built-in Power Query features.
7. **Build a function library** — stop solving the same problem twice.

---

## Q & A

**Thank you for attending!**

Slides, code samples & further reading:

- `linkedin.com/in/your-profile`
- `github.com/your-repo`
- `your-blog.com`

---

*© 2025 · Custom Functions in Power Query · All examples shared freely*
