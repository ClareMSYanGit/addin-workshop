# Lesson 2 answers


2.0.1 prep

```
function run() {
    Excel.run(async function (context) {
        var range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load(["address", "values"]);
        await context.sync()
        console.log("The range address was \"" + range.address + "\".");
        return populateRange(context, range);
    })
        .catch(function (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        });
}
```

2.5 Grand Total button

```
<button id="grand-total" class="ms-Button">
        <span class="ms-Button-label">Grand Total</span>
</button>

async function grandTotal() {
    try {
        await Excel.run(async (ctx) => {
            var range = ctx.workbook.worksheets.getItem("Sample").getRange("E3:E5");
            var rangeTot = ctx.workbook.worksheets.getItem("Sample").getRange("B7:E8");
            var gTot = ctx.workbook.functions.sum(range);

            range.load("values");
            rangeTot.load("values");
            gTot.load();

            await ctx.sync();

            var vTot = rangeTot.values;

            console.log(gTot.value);
            console.log(range);
            vTot[0][3] = gTot.value;
            vTot[0][0] = "Grand Total";
            vTot[0][1] = "=sum(c3:c5)";

            rangeTot.values = vTot;

            await ctx.sync();
        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```
