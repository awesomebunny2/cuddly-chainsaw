/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



// The initialize function must be run each time a new page is loaded
Office.initialize = () => {

  // console.log(theresponse);

  // See whether the start up behavior is on or off
  Office.addin.getStartupBehavior().then(function(response) {
    // If response is 'load', check the box
    if (response == "Load") {
      console.log("startupBehavior is set to Load");

      $("#chk-set").prop( "checked", true );

    } else { // Otherwise, uncheck the box
      console.log("startupBehavior is set to not Load");
      $("#chk-set").prop( "checked", false );
    }
    console.log(response);
  });

  // console.log(behavior);

  // Office.addin.setStartupBehavior(Office.StartupBehavior.load);


  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  // document.getElementById("run").onclick = run;

  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    sheet.onChanged.add(onChange); // <!---- Bound the event

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });

};

async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);

      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";

      /*
      1. Verify the column's name is Basket
      2. Get the value of the cell
      3. Get the sheet called (the value of the cell)
      4. Move the row from the table to the table on the sheet called (the value of the cell)
      */

      

    });
  });
}

/**
 * When the user checks the Load on Startup box
 */

$("#chk-set").on("change", function() {
    if (this.checked) {
      // Checked the box
      console.log("Checked. Turning on startupBehavior.");
      // Set startUpBehavior to 'load'
      Office.addin.setStartupBehavior(Office.StartupBehavior.load);

    } else {
      // Unchecked the box
      console.log("Unchecked. Turning off startupBehavior");
      // Set startUpBehavior to 'none'
      Office.addin.setStartupBehavior(Office.StartupBehavior.none);
    }
})


// async function run() {
//   try {
//     await Excel.run(async context => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }