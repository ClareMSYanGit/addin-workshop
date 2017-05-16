# Lesson 4: Wrapping up with some new APIs

There are some new functions added recently and we can make use of them. First, make sure you have the correct version of Office that supports the API requirement set.

## 4.0 Prep

### 4.0.1

For the calculate() function below, turn off the automatic formula calculation by going to:
File > Options > Formulas > select **Manual** under Workbook Calculation

To use the Range.Calculate() method and for calculate(), see https://dev.office.com/reference/add-ins/excel/range

For the ConditionalFormat object, see:

- <https://github.com/OfficeDev/office-js-docs/blob/ExcelJs\_OpenSpec/reference/excel/conditionalformatcollection.md>

- <https://github.com/OfficeDev/office-js-docs/blob/ExcelJs\_OpenSpec/reference/excel/conditionalformat.md>

In order to use these two new APIs, you need Excel API requirement set 1.6 to be supported by the host application. Let's not take any chances, but instead add a programmatic check in our code before executing these. Go here to read about requirement sets and how to check in code:
<https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets>

Now that you have the information, add a button to calculate our price table range B2:E7 (including our totals) and also to apply conditional formatting to the price table numbers in B3:E5.

## 4.1 Add the button to recalculate the range of the prices table

## 4.2 In the same button handler, add code to apply conditional formatting


Hints:

- It might be convenient to remove the stock multiply buttons from the sample code to make room for your new buttons.

# Lesson 4 Answers

4.0.1 programmatic check for API set 1.6

```
if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

4.1 Recalculate

```
async function recalculate() {
    try {
        await Excel.run(async (ctx) => {
            console.log("Recalculating price table");
            if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
                var rangeSelection = "B2:E5";
                var range = ctx.workbook.worksheets.getItem("Sample")
                    .getRange(rangeSelection);
                range.calculate();
                await ctx.sync();
                console.log("Done recalculating price table!");
            }
            else {
                console.log("Can't recalculate in this host");
            }

        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```

4.2 Add conditional formatting

```
async function recalculate() {
	try {
		await Excel.run(async (ctx) => {
				console.log("Recalculating price table");
				if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
				var rangeSelection = "B2:E5";
				var range = ctx.workbook.worksheets.getItem("Sample")
				.getRange(rangeSelection);
				range.calculate();
				await ctx.sync();

				var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
				conditionalFormat.iconOrNull.style = "YellowThreeArrows";
				await ctx.sync()
				console.log("Added new yellow three arrow icon set.");
				console.log("Done recalculating price table!");
				}
				else {
				console.log("Can't recalculate in this host");
				}

				});
	}
	catch (error) {
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
		OfficeHelpers.UI.notify(error);
		OfficeHelpers.Utilities.log(error);
	}
}
```