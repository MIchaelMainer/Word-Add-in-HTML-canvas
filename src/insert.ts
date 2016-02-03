

module ContextMenuButton {
    export function insertImage() { 
        // The initialize function is run each time the page is loaded.
        Office.initialize = (reason) => {
            // https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/66b137b536319294daa72b45888a46da906e2c81/Excel/Webapp/ODSampleDataWeb/Scripts/App/UX.js
            // https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/66b137b536319294daa72b45888a46da906e2c81/Excel/Webapp/ODSampleDataWeb/Scripts/App/UX.ts
            // https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/66b137b536319294daa72b45888a46da906e2c81/Excel/Webapp/ODSampleDataWeb/Scripts/App/App.ts
            
                    Word.run(function (context) {
                    // Create a proxy object for the range at the current selection.
                    var imageRange = context.document.getSelection();
                            
                    // Load the selected range.
                    context.load(imageRange, 'text');
                
                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync()
                        .then(function () {
                            // Queue a command to insert the image into the document.
                            var insertedImage = imageRange.insertText('This is text', Word.InsertLocation.replace);
                                    
                            // Queue a command to navigate the UI to the insert picture.
                            insertedImage.select();
                        })
                    // Synchronize the document state by executing the queued commands.
                        .then(context.sync);
                })
                    .catch(function(error) {
                        console.log('Error: ' + JSON.stringify(error));
                        if (error instanceof OfficeExtension.Error) {
                            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                        }
                    });
            
        };
    }
}
        
//             // 1. This UILess function can be triggered by 'GetData' button (id=Contoso.Button1Id1) or context menu 'GetData' button (id=Contoso.TestMenu1)
//     // 2. The first clicking for any UILess function bound ribbon button or context menu item triggers office's initialize() firstly. Other clickings including other buttons don't trigger the initialize() again.
//     // 3. The UILess processing is invoked when the user clicks the bound ribbon button or context menu item 
//     // 4. args.completed() is needed at the time the UILess processing is completed, otherwise the other UILess processing would not be invoked until 300s timeout.
//     // 5. These UILess processings include all other buttons/contextmenus bound processings.
//     export function getButton(args) {
//         if (buttonId != 0) {
//             buttonId = 1;
//             ODSampleData.onOfficeReady();
//         }
//         buttonId = 1;

//         args.completed();
//     }
// }