# Week 1 Interview Practice Questions

**Excel Foundations & Analytical Skills** – organized by topic, with model answers

---

## 1. Navigating & Formatting Excel

### Q1. What's the difference between Ctrl+End and Ctrl+Arrow, and when would you use each?

**Model Answer:** Ctrl+End jumps to the last used cell in the entire sheet; Ctrl+Arrow jumps to the edge of the current contiguous data block in whichever direction you press. Use Ctrl+End to find the true extent of a sheet's data, and Ctrl+Arrow to move quickly within or between data blocks without scrolling.

### Q2. A column of dates displays as five-digit numbers like 46047 instead of readable dates. What's actually happening, and how do you fix it?

**Model Answer:** The cells already hold real date values, but they're formatted as General or Number instead of Date, so Excel is showing the underlying serial number. Select the column and apply a Date format – this fixes the display without changing the underlying value at all.

### Q3. Why would you freeze panes on a large dataset, and what's the difference between freezing just row 1 versus freezing row 1 and column A?

**Model Answer:** Freezing keeps key reference cells visible while scrolling through a lot of data. Freezing row 1 alone keeps column headers visible while scrolling down; freezing row 1 and column A also keeps a key identifying column (like an ID) visible while scrolling sideways as well.

---

## 2. Organizing & Structuring Data

### Q4. What's the practical benefit of converting a range into a Table (Ctrl+T) instead of leaving it as a plain range?

**Model Answer:** Tables auto-expand as new rows or columns are added, apply consistent banded formatting automatically, support structured references in formulas, and integrate cleanly with pivot tables, sorting, and filtering.

### Q5. You see a formula like `[@[Hours Logged]]` inside a Table. What does the @ symbol mean, and why use it instead of a plain cell reference?

**Model Answer:** The @ means "this row" – the formula always refers to that column's value in the same row it's written in. Unlike a plain relative reference, it's inherently row-relative and self-documenting, so it doesn't need to be manually adjusted as it's filled down.

### Q6. You want to filter a dataset to only Branch = 'Downtown' AND Status = 'Escalated' at the same time. How do you set that up, and what would change if you needed an OR condition instead?

**Model Answer:** Apply the AutoFilter dropdown on each column separately and check only the desired value in each – column filters combine as AND automatically. A true OR condition across two different columns isn't something a simple AutoFilter can do on its own; that typically needs a helper column or Advanced Filter instead.

---

## 3. Core Formulas & Cell References

### Q7. What's the difference between COUNT and COUNTA, and give an example of when using the wrong one would give a misleading result?

**Model Answer:** COUNT only counts cells containing numbers; COUNTA counts any non-blank cell regardless of type. Running COUNT on a column of text-based Status values would return 0 even though every row is filled in – COUNTA would correctly count them.

### Q8. You write `=A2*$B$2` and copy it down column C. Walk through what happens to each part of the reference as it's copied.

**Model Answer:** `A2` is relative and shifts to `A3`, `A4`, and so on as the formula is copied down each row. `$B$2` is absolute (locked with dollar signs) and stays fixed at B2 in every copied row – useful when multiplying every row by one shared constant, like a single tax rate.

### Q9. When would MEDIAN give a more useful picture of a dataset than AVERAGE?

**Model Answer:** When the data has outliers or is skewed. A few unusually large or small values pull AVERAGE away from what's typical, while MEDIAN reflects the actual middle value and isn't distorted by extremes.

---

## 4. Percentage & Financial Calculations

### Q10. Walk through how you'd calculate Variance % between an actual and a budgeted value, and explain what a negative result means.

**Model Answer:** Variance % = (Actual − Budget) / Budget. A negative result means the actual value came in under the budgeted or expected amount; a positive result means it exceeded it.

### Q11. A transaction shows a Margin % of −12%. Is that necessarily a broken formula? Explain.

**Model Answer:** Not necessarily – a negative margin simply means cost exceeded revenue on that transaction, producing a loss. That's a legitimate real-world outcome the formula is correctly reporting, not automatically a sign of an error.

### Q12. A trainee calculates variance as `=(Actual-Budget)/Actual` instead of dividing by Budget. What's conceptually wrong with this, even though it still produces a number?

**Model Answer:** Dividing by Actual instead of Budget changes what the percentage represents. Variance is meant to measure how far the outcome is from the original expectation (Budget), not relative to itself – using the wrong denominator produces a number that calculates fine but doesn't actually answer the intended business question.

---

## 5. Errors, Validation & Logical Checks

### Q13. What's the difference between #N/A and #VALUE!, and what's a common cause of each?

**Model Answer:** #N/A usually comes from a lookup, like VLOOKUP, that couldn't find a matching value. #VALUE! usually comes from trying to perform math on text that isn't actually numeric, such as adding a number to a text string.

### Q14. Why wrap a VLOOKUP in IFERROR rather than leaving it as-is, and what's the tradeoff of doing so?

**Model Answer:** IFERROR replaces any error, like a #N/A, with a cleaner fallback value (often a blank), which keeps the sheet readable and prevents the error from breaking downstream formulas. The tradeoff is that it can also quietly hide a genuine bug in the lookup itself if you stop investigating once the error disappears.

### Q15. You need to confirm that no row in a 50-row tracker was accidentally left without a Match Status. What formula would you use, and what result tells you everything is fine?

**Model Answer:** `=COUNTBLANK()` over the Match Status range. A result of 0 confirms every single row has a value; anything higher means at least one row was skipped.

---

## 6. Cleaning Messy Text

### Q16. What's the difference between TRIM and CLEAN, and why might a cell need both?

**Model Answer:** TRIM removes leading, trailing, and extra internal regular spaces. CLEAN removes non-printing characters, like a stray line break, that aren't regular spaces and that TRIM can't touch. A cell can still show an inflated `LEN()` after TRIM if a hidden character like this is present, which is when CLEAN is also needed.

### Q17. Two Client Name values look identical at a glance, but a formula comparing them returns FALSE. What are two possible explanations, and how would you confirm which one it is?

**Model Answer:** Most likely a hidden difference invisible to the eye – a trailing space, a non-printing character, or inconsistent casing if using EXACT specifically. Compare `LEN()` on both cells to spot a length mismatch, or use `EXACT()` to confirm whether casing is actually the issue, to pin down the real cause before fixing it.

---

## 7. Parsing & Combining Text

### Q18. How would you extract just the numeric portion from an ID like 'ENG-4032' using text functions, without using Text to Columns?

**Model Answer:** Use FIND to locate the position of the dash, then MID (or RIGHT) to pull the characters after it – for example, `=MID(A5,FIND("-",A5)+1,10)`.

### Q19. You want to rebuild a 'City, State' value from two separately cleaned columns, skipping either one if it's blank. Which function fits best, and why?

**Model Answer:** TEXTJOIN, because it lets you specify a delimiter and can automatically skip blank cells – CONCATENATE would just glue the values together and leave an awkward stray delimiter if one side were blank.

### Q20. What does `LEN()` tell you that's useful when cleaning text, beyond just a character count?

**Model Answer:** It's a fast way to catch inconsistencies that aren't visible on screen, like a hidden non-printing character or an extra space, by comparing the length of a value before and after a cleaning formula – if two visually identical values have different `LEN()` results, something is hiding in one of them.

---

## 8. Restructuring & Converting Data

### Q21. A column of dates is left-aligned and won't sort correctly. What's actually wrong, and what are two different ways to fix it?

**Model Answer:** The dates are stored as text, not real date values – the left-alignment confirms it. Fix with Text to Columns (Delimited, then Finish with Date chosen as the column format) if the pattern is consistent across the column, or with a DATEVALUE formula wrapped in IFERROR if formats vary row to row.

### Q22. Why is it risky to run Text to Columns on a column without first checking what's in the columns immediately to its right?

**Model Answer:** Text to Columns overwrites whatever is already in the destination columns by default, without a separate warning about existing content – if those columns already hold other data, running it will silently overwrite or delete it.

### Q23. How can you convert a column of numbers stored as text into real numbers without writing a single formula?

**Model Answer:** Select the column, go to **Data > Text to Columns > Delimited**, leave every delimiter unchecked, and click Finish – Excel re-parses the values as real numbers on the way back in.

---

## 9. Data Integrity Tools

### Q24. If you select two columns at once and apply Conditional Formatting > Highlight Duplicate Values, what actually gets compared?

**Model Answer:** The entire selected range is treated as one shared pool of values – a value in one column can get flagged as a duplicate of a value in the other column, which usually isn't the intent. To check each column independently, the rule needs to be applied to one column at a time.

### Q25. A trainee runs Remove Duplicates with every column checked when they only meant to de-duplicate by Transaction ID. What went wrong, and how should it have been set up?

**Model Answer:** Checking every column requires an exact match across all of them to count as a duplicate, which can miss real duplicates that share the same ID but differ in some other field (due to a later edit, for example). Only the column(s) that actually define uniqueness – here, just the ID – should be checked.

### Q26. What's the main limitation of a Data Validation dropdown list as a data-quality tool?

**Model Answer:** It only restricts what can be typed in going forward – it does nothing to fix, flag, or even detect values that were already sitting in the cells before the rule was applied.

---

## 10. Lookups & Reconciliation

### Q27. Why is the final FALSE argument in VLOOKUP important, and what could go wrong if it were omitted or set to TRUE?

**Model Answer:** FALSE requires an exact match. Without it, VLOOKUP performs an approximate match against a sorted list, which can silently return the wrong row's value instead of failing with a visible error – a much more dangerous failure mode than an obvious #N/A.

### Q28. You're reconciling two systems, and a VLOOKUP against a supposedly-matching ID keeps returning #N/A even though you can see the transaction in both sheets. What's the most likely cause, and how would you confirm it?

**Model Answer:** The ID formatting most likely differs between the two systems – casing, dashes versus underscores, or stray spaces – so the values don't match exactly even though they look the same. Confirm with `EXACT()` or by comparing `LEN()` on both IDs, then fix it by standardizing both sides to the same format before looking up.

### Q29. Why build a master list of unique keys from both systems, by stacking IDs and running Remove Duplicates, before starting a reconciliation, rather than just using one system's list as-is?

**Model Answer:** A single system's list won't include transactions that exist only in the other system. Building the union of both is the only way to guarantee the reconciliation catches every "missing from one side" case, not just amount mismatches on records both systems already share.

---

## 11. Pivot Tables & Reporting

### Q30. What's the difference between using the Filters area and using a Slicer to filter a pivot table?

**Model Answer:** The Filters area applies a dropdown filter to a single pivot table. Slicers give the same filtering as clickable buttons, are faster to use interactively, and can be connected to multiple pivot tables at once so several related reports filter together.

### Q31. You add a field to the Filters area of a pivot table built starting at the very top of a worksheet, and it produces a #SPILL error. What's actually going on?

**Model Answer:** Adding a Filters field requires Excel to insert an extra row above the pivot table's top-left corner; if the table starts in row 1, there's no room to grow upward, and current Excel versions surface that failure as a #SPILL error. The fix is to move or rebuild the pivot table a few rows lower, leaving blank space above it.

### Q32. Why would you group a Date field by Month in a pivot table instead of leaving every individual date as its own row?

**Model Answer:** Grouping collapses many individual dates into meaningful reporting periods, turning an unreadable list of one-off dates into a clean monthly (or quarterly) summary that's actually useful for a manager or stakeholder to review.

---

## 12. Productivity Habits

### Q33. What does Paste Special > Values Only actually do, and why use it after a formula has already calculated a result?

**Model Answer:** It pastes only the calculated value, not the underlying formula, effectively freezing the number in place. This is useful when you want to lock in a result before deleting the source data the formula depended on, or before sharing a static snapshot that shouldn't recalculate later.

### Q34. An associate doesn't know whether a particular Excel function exists for what they're trying to do. What's the recommended way to find out, rather than guessing or giving up?

**Model Answer:** Use the Insert Function (fx) button next to the formula bar, or browse the Formulas ribbon tab, and search by keyword – this is the standard way to discover any Excel function without needing it explicitly taught beforehand.

### Q35. What does the Fill Handle actually do when you drag a formula down a column, and how does it interact with relative versus absolute references?

**Model Answer:** It copies the formula into each new cell, automatically adjusting relative references (like `A2`) to match the new row while keeping absolute references (like `$A$2`) fixed in place – this combination is what makes it fast to apply the same calculation logic across many rows at once.
