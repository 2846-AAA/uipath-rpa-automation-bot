# Phase 2 — UiPath Studio Web Form Automation Guide

## What the Python script simulates

The `phase2_webform_bot.py` script **mirrors** exactly what a UiPath `.xaml` workflow does. Here's the mapping:

---

## UiPath Workflow Structure (phase2_webform.xaml)

```
Main Sequence
│
├── 1. Read JSON File Activity
│   └── Deserialize JSON → employees_clean.json
│
├── 2. For Each Record (ForEach activity)
│   │
│   ├── Open Browser (Chrome/Edge)
│   │   └── URL: "https://your-form-url.com"
│   │
│   ├── TypeInto — Employee ID field
│   │   └── Selector: <input id='emp_id' />
│   │
│   ├── TypeInto — Full Name field
│   │   └── Selector: <input id='full_name' />
│   │
│   ├── TypeInto — Department
│   ├── TypeInto — Salary
│   ├── TypeInto — Join Date
│   ├── TypeInto — Email
│   │
│   ├── Click — Submit button
│   │   └── Selector: <button type='submit' />
│   │
│   ├── GetText — Read confirmation message
│   │
│   └── Log Message — Record result
│
├── 3. Error Handling (Try/Catch)
│   └── On exception → Log + Retry Scope (2 retries)
│
└── 4. Write Results to JSON
```

---

## Selector Management Tips

```xml
<!-- Reliable selector — uses stable ID attribute -->
<webctrl id='emp_id' tag='INPUT' />

<!-- Fallback if ID changes — use multiple attributes -->
<webctrl name='employee_id' tag='INPUT' type='text' />

<!-- For dynamic pages — combine CSS class + position -->
<webctrl css-selector='.form-field:nth-child(1) input' />
```

## Error Handling Pattern

```
Try
  └── TypeInto activities + Click Submit
Catch (SelectorNotFoundException)
  └── Log "Selector not found: " + exception.Message
  └── Take Screenshot for debugging
  └── Continue (skip to next record)
Catch (TimeoutException)
  └── Log "Timeout waiting for element"
  └── Retry Scope → max 2 retries, delay 5s
```

---

## Running in UiPath Studio

1. Open UiPath Studio Community Edition
2. Create new **Process** project
3. Drag activities from the Activities panel:
   - `Read Text File` → read employees_clean.json
   - `Deserialize JSON` → parse to JArray
   - `For Each` → iterate records
   - `Open Browser` → launch Chrome
   - `TypeInto` × 6 → fill form fields
   - `Click` → submit
   - `Get Text` → capture confirmation
4. Add `Try Catch` around the inner sequence
5. Add `Log Message` activities throughout
6. Run with `Ctrl+F5` to see live execution
