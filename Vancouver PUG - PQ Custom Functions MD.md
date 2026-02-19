
# Getting Crafty with Power Query Custom Functions in Power BI

> Build reusable, maintainable, and powerful data transformations
> 
> All levels welcome

&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
&nbsp;
       
&nbsp;
&nbsp;
&nbsp;
&nbsp;
       
&nbsp;
&nbsp;
&nbsp;
&nbsp;
       
&nbsp;

## About me   
&nbsp;

# Cristian Angyal
<img width="233" height="384" alt="image" src="https://github.com/user-attachments/assets/46242c81-10d7-4ea3-b451-9610aade0d9a" />

- **Project Manager** (PMP, PMI-ACP)
- **Microsoft Certified Professional** (DP-600, PL-300, MCT, MCSA, MCSE, MOS Master)
- **Romania PBI and Modern Excel UG Founder** — [https://www.meetup.com/romaniapug](https://www.meetup.com/romaniapug)
- **Microsoft MVP** M365 (**Excel**) & Data Platform (**Power BI**)
- **Proud dad and husband** 

🔗 [linkedin.com/in/cristian-angyal](https://linkedin.com/in/cristian-angyal)  
🔗 [https://linktr.ee/cristianangyal](https://linktr.ee/cristianangyal)

---


&nbsp;

## Agenda

| # | Section |
|---|---------|
| 01 | Why Custom Functions? |
| 02 | Syntax & Structure |
| 03 | M Keywords Explained |
| 04 | Parameters & Types |
| 05 | Demos |
| 06 | Error Handling |
| 07 | Performance Tips |
| 08 | Reusing Your Functions |
| 09 | Q&A & Wrap-Up |

&nbsp;

---

&nbsp;

## 01 · Why Custom Functions?

*From copy-paste chaos to reusable, elegant M code*
> *We're not lazy, we're efficient :)*

&nbsp;

❌ **Duplicate logic** - Same transformation repeated in 20 queries. One bug → 20 fixes.

❌ **Hard to maintain** - Next month you won't remember what that nested `List.Transform` was doing.

❌ **Not testable** - No easy way to unit-test an inline transformation step.

&nbsp;

✅ **Custom Functions solve all of this** 
- Write once, call everywhere. Self-documenting. Parameterized.

&nbsp;

---

&nbsp;

## 02 · Syntax & Structure

*The anatomy of an M function*

&nbsp;

```m
(parameter1 as type, parameter2 as type)
    as returnType =>
let
    step1 = ...,
    step2 = ...
in
    step2
```

&nbsp;

① **Parameters list** - just like any programming language

② **Return type** - optional, but recommended for documentation

③ **`let…in` block** - exactly the same as a regular query!

④ **Final expression** after `in` is the return value - no `return` keyword needed

&nbsp;
  
# Refining function definitions

* **Defining data types:** => ensures the data remains consistent and accurate so that the function input is as expected.
* **Adding optional parameters:** => allows users flexibility in how the function is used, addressing diverse needs without rigid constraints.

For both input parameters and function outputs, you can define data types. 

A parameter without a defined data type is referred to as an **implicit parameter**. This means it can accept values of any type. 

Specifying a data type for a parameter turns it into an **explicit parameter**. Once you define a data type for your input parameter, the custom functions only accept values of the specified type.   

---

&nbsp;

### 3 Ways to Create a Custom Function

&nbsp;

**Method 1 - Blank Query → Paste M code**

Most flexible. Full control. Best for complex, reusable functions.

```m
//Example function: adds 10 days to the input date
   (date as date) as date => Date.AddDays(date, 10)
```

**Method 2 - Convert an existing Query to a Function**

Right-click any query → *Create Function*. Power Query wraps it automatically.

Three main advantages:
* **Ease of creating functions**
* **Function updates automatically** 
* **Visibility of the query**: The original query remains visible in your project. That means you can still inspect, modify, and run your logic as a standalone query.

**Method 3 - Define inline inside another query**

Define and call immediately - great for one-off local helpers.

```m
let
    Double = (x as number) => x * 2,
    Result = Double(5)          // = 10
in
    Result
```

&nbsp;
&nbsp;
&nbsp;
---

&nbsp;

## 03 · M Keywords Explained

*Every keyword you will use in a custom function*

&nbsp;

&nbsp;

### `let` … `in`

Defines a sequence of named steps. Each step can reference the ones above it.
The expression after `in` is the **output** of the entire block.

```m
let
    step1 = 10,
    step2 = step1 * 2,   // = 20
    step3 = step2 + 5    // = 25
in
    step3                // returns 25
```

> `let…in` works identically inside a function body and inside a regular query.

&nbsp;

---

&nbsp;

### `=>` (Fat Arrow, Goes To, Rocket Sign)

Separates the **parameter list** from the **function body**.
Think of it as "given these inputs, produce this output."

```m
// Minimal: no parameters, returns a constant
() => 42

// With one parameter
(x as number) => x * x

// With a let…in body
(x as number) =>
let
    doubled = x * 2,
    msg     = "Result: " & Number.ToText(doubled)
in
    msg
```

&nbsp;

---

&nbsp;

### `as`

Declares a **type** for a parameter or the function return value.
Optional, but strongly recommended - documents intent and enables the UI parameter dialog.

```m
// Parameter type            Return type
(name as text)               as text =>
    "Hello, " & name & "!"
```

| Position | Meaning |
|----------|---------|
| `(x as number)` | x must be a number |
| `as text =>` | this function returns text |
| `as nullable date` | date or null is accepted |
| `as optional text` | caller may omit this argument |

&nbsp;

---

&nbsp;

### `nullable`

Allows a parameter to accept **either its declared type or `null`**.
Essential for real-world data - table columns almost always contain nulls.

```m
// Without nullable → crashes on null input
(code as text) as text => Text.Upper(code)

// With nullable → safe
(code as nullable text) as text =>
    if code = null then "" else Text.Upper(code)
```

> Always use `nullable` for parameters applied to table columns.

&nbsp;

---

&nbsp;

### `optional`

Marks a parameter the **caller can omit entirely**.
Inside the function the omitted value arrives as `null` - pair it with `??`.

```m
(name as text, greeting as optional text) as text =>
let
    g = greeting ?? "Hello"    // default when omitted
in
    g & ", " & name & "!"

// With greeting:   fn("Ana", "Bună ziua")  →  "Bună ziua, Ana!"
// Without:         fn("Ana")               →  "Hello, Ana!"
```

&nbsp;

---

&nbsp;

### `??` (Null Coalescing Operator)

Returns the **left side** if it is not null; otherwise returns the **right side**.
The idiomatic way to provide default values.

```m
let
    userInput = null,
    result    = userInput ?? "default value"   // = "default value"
in
    result

// Typical use in a function:
(label as optional text) =>
    let display = label ?? "N/A"
    in  display
```

&nbsp;

---

&nbsp;

### `each` and `_`

`each` is shorthand for a single-parameter anonymous function.
`_` is the implicit parameter name inside an `each` expression.

```m
// These two expressions are identical:
List.Transform({1, 2, 3}, each _ * 2)
List.Transform({1, 2, 3}, (x) => x * 2)

// Common patterns:
Table.SelectRows(Source, each [Status] = "Active")
List.Select({1..10},     each _ > 5)
```

> `each` is syntax sugar - use it for short, readable one-liners.

&nbsp;

---

&nbsp;

### `try` … `otherwise`

Evaluates an expression and **catches any error**, returning a fallback value instead of crashing.

```m
// Without try → crashes if text is not a valid number
Number.FromText("abc")                    // Error!

// With try…otherwise → safe fallback
try Number.FromText("abc") otherwise 0   // = 0
try Number.FromText("42")  otherwise 0   // = 42
```

The full `try` result is a **record** you can inspect:

```m
let r = try Number.FromText("abc")
// r = [ HasError = true, Error = [Reason = ..., Message = ...] ]
in  if r[HasError] then -1 else r[Value]
```

&nbsp;

---

&nbsp;

### `error`

**Raises** a custom error - validate inputs and give callers a meaningful message.

```m
(rate as number) as number =>
    if rate < 0 or rate > 1
    then error Error.Record(
             "InvalidRate",
             "Rate must be between 0 and 1",
             rate)
    else rate

// Error.Record(Reason, Message, Detail)
```

> Combine `error` for raising with `try…otherwise` for catching.

&nbsp;

---

&nbsp;

### `#shared`

A built-in record containing **every function and query currently loaded** in the session.
Used when dynamically evaluating M code at runtime.

```m
// Evaluate M code from a text string - needs #shared as the environment
Expression.Evaluate("Text.Upper(""hello"")", #shared)
// = "HELLO"
```

> You will encounter `#shared` mainly when building a local `.txt` function library (Section 08).

&nbsp;

---

&nbsp;

## 04 · Parameters & Types

*Make your functions safe and self-documenting*

&nbsp;

### M Primitive Types

| Type | Example |
|------|----------------|
| `text` | `"Hello"` |
| `number` | `42`, `3.14` |
| `logical` | `true`, `false` |
| `date` | `#date(2025, 1, 1)` |
| `datetime` | `#datetime(2025, 1, 1, 9, 0, 0)` |
| `duration` | `#duration(1, 0, 0, 0)` - 1 day |
| `list` | `{1, 2, 3}` |
| `record` | `[Name = "Ana", Age = 30]` |
| `table` | any Power Query table |
| `function` | a function value (yes, functions are values!) |
| `nullable X` | X or null |

&nbsp;

---

&nbsp;

### Self-Documenting Functions with Metadata

Power Query shows a rich parameter dialog when you add `Documentation` metadata.

```m
let
    fnBody = (n as number, optional factor as number) as number =>
    let
        f      = factor ?? 1,
        result = n * f
    in  result,

    fnType = type function
        (n as (type number meta [
            Documentation.FieldCaption = "Base Number",
            Documentation.SampleValues = {10, 100}
         ]),
         optional factor as (type number meta [
            Documentation.FieldCaption     = "Multiplier",
            Documentation.FieldDescription = "Defaults to 1 if omitted",
            Documentation.SampleValues     = {2, 10}
         ])) as number
    meta [
        Documentation.Name            = "fnMultiply",
        Documentation.LongDescription = "Multiplies **n** by an optional **factor** (defaults to 1).",
        Documentation.Examples        = {
            [
                Description = "Multiply 10 by 3",
                Code        = "fnMultiply(10, 3)",
                Result      = "30"
            ],
            [
                Description = "Without factor (defaults to 1)",
                Code        = "fnMultiply(10)",
                Result      = "10"
            ]
        }
    ],

    documented = Value.ReplaceType(fnBody, fnType)
in  documented
```

&nbsp;

---

&nbsp;

## 05 · Demos

*Self-contained demos - paste and run*

&nbsp;

---

&nbsp;

### Demo 1 - Greeting with a Default

Shows **parameters, `optional`, and `??`** in a fully self-contained query.

```m
// Paste this as a single blank query and run it
let
    fn_Greet = (name as text, optional language as text) as text =>
    let
        lang     = language ?? "EN",
        greeting = if      lang = "RO" then "Bună ziua, "
                   else if lang = "FR" then "Bonjour, "
                   else                     "Hello, "
    in
        greeting & name & "!",

    Results = {
        fn_Greet("Ana"),            // → "Hello, Ana!"
        fn_Greet("Ion",   "RO"),    // → "Bună ziua, Ion!"
        fn_Greet("Marie", "FR")     // → "Bonjour, Marie!"
    }
in
    Results
```

&nbsp;

---

&nbsp;

### Demo 2 - Apply a Function Across a Table

Shows **`Table.AddColumn`** and **`each`** with inline sample data.

```m
// Paste this as a single blank query and run it
let
   Source = #table(
        {"Product",   "Price"},
        {{"Keyboard",  250  },
         {"Mouse",      89  },
         {"Monitor",  1200  },
         {"USB Hub",    45  }}
    ),
   fn_TaxAmount = (price as number, optional rate as number) as number =>
        Number.Round(price * (rate ?? 0.19), 2),

    WithTax   = Table.AddColumn(Source, "VAT",   each fn_TaxAmount([Price]),  type number),
    WithTotal = Table.AddColumn(WithTax, "Total", each [Price] + [VAT],       type number)
in
    WithTotal
```

| Product  | Price |   VAT |  Total |
| -------- | ----: | ----: | -----: |
| Keyboard |   250 |  47.5 |  297.5 |
| Mouse    |    89 | 16.91 | 105.91 |
| Monitor  |  1200 |   228 |   1428 |
| USB Hub  |    45 |  8.55 |  53.55 |

&nbsp;

---

&nbsp;

### Demo 3 - Repeat a Function Over a List

Shows **`List.Transform`** - Power Query's for-loop.

```m
// Paste this as a single blank query and run it
let
    CelsiusValues    = { 0, 10, 20, 37, 100},
    
    fn_ToFahrenheit = (c as number) as number =>
        Number.Round(c * 9 / 5 + 32, 1),
    FahrenheitValues = List.Transform(CelsiusValues, fn_ToFahrenheit),

    Result = Table.FromColumns(
        {CelsiusValues, FahrenheitValues},
        {"°C", "°F"}
    )
in
    Result
```

|  °C |   °F |
| --: | ---: |
|   0 |   32 |
|  10 |   50 |
|  20 |   68 |
|  37 | 98.6 |
| 100 |  212 |

> 💡 Replace `fn_ToFahrenheit` with any function - the pattern is always the same.

&nbsp;

---

&nbsp;

## 06 · Error Handling

*Write functions that fail gracefully*

&nbsp;


&nbsp;

### Pattern 1 - `try … otherwise`

Returns a fallback value if any error occurs. The simplest safety net.

```m
(input as text) as number =>
    try Number.FromText(input) otherwise 0

// fn("42")   → 42
// fn("abc")  → 0   (no crash)
// fn(null)   → 0   (no crash)
```

&nbsp;

---

&nbsp;

### Pattern 2 - Null Guard with `??`

Always handle null before doing anything with the value.

```m
(code as nullable text) as text =>
let
    safe   = code ?? "",
    result = Text.Upper(Text.Trim(safe))
in
    result

// fn("hello ")  → "HELLO"
// fn(null)      → ""   (no crash)
```

&nbsp;

---

&nbsp;

### Pattern 3 - `try` with Error Record

Inspect *what went wrong* - useful for logging bad rows instead of dropping them.

```m
(input as text) as text =>
let
    r = try Number.FromText(input)
in
    if r[HasError]
    then "ERROR: " & r[Error][Message]
    else Number.ToText(r[Value])

// fn("42")   → "42"
// fn("abc")  → "ERROR: ..."
```

&nbsp;

---

&nbsp;

### Pattern 4 - Raise a Custom Error

Validate inputs and give callers a meaningful failure message.

```m
(rate as number) as number =>
    if rate < 0 or rate > 1
    then error Error.Record(
             "InvalidRate",
             "Rate must be between 0 and 1",
             rate)
    else Number.Round(rate * 100, 2)

// fn(0.19)  → 19
// fn(1.5)   → Error: "Rate must be between 0 and 1"
```

&nbsp;

> ⚠️ **A function without null handling WILL break in production.**

&nbsp;

---

&nbsp;

## 07 · Performance Tips

*Write functions that scale - not just work*

&nbsp;

⚡ **Avoid row-level function calls on large tables**
Each `Table.AddColumn` with a custom function is called once per row. 1M rows = 1M M evaluations. Prefer built-in column operations where possible.

&nbsp;

🔁 **Protect Query Folding**
If your source supports folding (SQL Server, SharePoint…), custom M steps break it. Check: *right-click a step → View Native Query*. Greyed out = folding broken.

&nbsp;

📦 **Cache expensive calls outside the function**
Never call `Web.Contents()` or `File.Contents()` inside a function applied per row.

```m
// ❌ Wrong - file read happens once per row
WithData = Table.AddColumn(Source, "Rate", each fn_LoadRateFromFile([Currency]))

// ✅ Right - read once, pass the result in
RateTable = fn_LoadRateFromFile("all"),
WithData  = Table.AddColumn(Source, "Rate",
                each Table.SelectRows(RateTable, each [CCY] = _){0}[Rate])
```

&nbsp;

🧪 **Develop with `Table.FirstN`**
Wrap your source in `Table.FirstN(Source, 100)` while building and testing.

&nbsp;

---

&nbsp;

## 08 · Reusing Your Custom Functions

*From this workbook to every workbook*

&nbsp;

&nbsp;

### Method 1 - Copy & Paste the Query

The fastest way. Works in both Power BI Desktop and Excel.

1. In Power Query Editor, **right-click** the function in the Queries pane
2. Select **Copy**
3. Open the destination file → Power Query Editor
4. **Right-click** in the Queries pane → **Paste**

&nbsp;

✅ Instant - no setup required
⚠️ Each copy is independent - changes to the original do not propagate

&nbsp;

---

&nbsp;

### Method 2 - Share the M Code as Text

Extract the raw M code and share it via email, Teams, a wiki, or a GitHub Gist.

**Extract:** Open the function query → **Home → Advanced Editor** → Select All → Copy

**Import:** New Blank Query → **Advanced Editor** → Paste → Done

```m

(n as number, optional factor as number) as number =>
let
    f = factor ?? 1
in
    Number.Round(n * f, 2)
```

&nbsp;

✅ Easy to share across teams and platforms
✅ Version-controllable in Git
⚠️ Manual - each recipient pastes and maintains their own copy

&nbsp;

---

&nbsp;

### Method 3 - Local `.txt` Function Library

Store your functions as plain M code in `.txt` files.
Load them at runtime with `File.Contents` + `Expression.Evaluate`.

**Folder structure:**

```
📁 C:\PQ_FunctionLibrary\
    ├── fn_Greet.txt
    ├── fn_TaxAmount.txt
    └── fn_ToFahrenheit.txt
```

**Content of `fn_TaxAmount.txt`:**

```m
(price as number, optional rate as number) as number =>
let
    r = rate ?? 0.19
in
    Number.Round(price * r, 2)
```

**Load and call from any Power Query session:**

```m
let
    fn_TaxAmount = Expression.Evaluate(
                       Text.FromBinary(
                           File.Contents("C:\PQ_FunctionLibrary\fn_TaxAmount.txt")),
                       #shared),

    Source   = #table({"Product","Price"}, {{"Keyboard",250},{"Mouse",89}}),
    WithTax  = Table.AddColumn(Source, "VAT", each fn_TaxAmount([Price]), type number)
in
    WithTax
```

&nbsp;

✅ **Single source of truth** - update the `.txt` file once, every report benefits on next refresh
✅ Works in any Power BI or Excel file that can reach the folder
✅ No SharePoint or external services needed
⚠️ Requires `Expression.Evaluate` → enable in Power BI: **File → Options → Security → Allow user-defined M functions**
⚠️ Use a shared network path so the whole team can reach the library

&nbsp;

---

&nbsp;

### Comparison

| | Copy & Paste | Share M Code | `.txt` Library |
|--|:--:|:--:|:--:|
| Setup effort | 🟢 None | 🟢 Minimal | 🟡 Some |
| Stays in sync | ❌ No | ❌ No | ✅ Yes |
| Team sharing | 🟡 Manual | 🟡 Manual | 🟢 Shared folder |
| Version control | ❌ No | 🟢 Via Git | 🟢 Via Git |
| Best for | Quick one-off reuse | Sharing snippets | Team function library |

&nbsp;

---

&nbsp;

## Key Takeaways

&nbsp;

1. Functions use the same `let…in` syntax as queries - nothing new to learn
2. Always type your parameters and guard against nulls - safety first
3. Return a **record** to produce multiple columns from a single function call
4. `Table.AddColumn` + `each` and `List.Transform` are your core application tools
5. Test with `Table.FirstN` during development; protect query folding in production
6. Add `Documentation` metadata to make functions feel like built-ins
7. **Build a `.txt` function library** - stop solving the same problem twice

&nbsp;

---

&nbsp;

## Q & A

**Thank you for attending!**

&nbsp;

Slides, code samples & further reading:

- `linkedin.com/in/your-profile`
- `github.com/your-repo`
- `your-blog.com`

&nbsp;

---
