.. _glossary:

Glossary
========

  model

    A subset of a workbook ("spreadsheet") comprising a collection of cells and the relationship between said cells. Usually identified by input cells and output cells.

    Example in Sheet1!B2;

    = Sheet1!C2 + Sheet1!D2

    In the above we can see the input cells are C2 and D2 (for they define inputs) and output cell B2 (as it defines the output). If this example were loaded as a model only the cells B2, C2 and D2 would be considered - the rest of the workbook would be ignored.

    It is obvious that as equations get more complex there can be more input cells. Counterintuitively there can also be more output cells. The reason for this is when you have equations referencing other equations there might be use in having the intermediate cells evaluate.
