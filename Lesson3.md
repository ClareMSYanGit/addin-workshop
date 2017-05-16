# Lesson 3: Charts

Now let's add a chart based on the code we've started in Lesson 2.

3.1 Add a button and code to chart the price table B2:E5.


Hints:

- Use the Chart Collection object, add method.

- See <https://dev.office.com/reference/add-ins/excel/chartcollection>

Answers
-------
```
async function createChart() {
    try {
        await Excel.run(async (ctx) => {
            var rangeSelection = "B2:E5";
            var range = ctx.workbook.worksheets.getItem("Sample")
                .getRange(rangeSelection);
            var chart = ctx.workbook.worksheets.getItem("Sample")
                .charts.add("ColumnClustered", range, "auto");
            await ctx.sync();
            console.log("New Chart Added");
        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```