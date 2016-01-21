

// module ContextMenuButton { 
//     export function insertImage(args) { 
//          // The initialize function is run each time the page is loaded.
//         Office.initialize = function (reason) { };
//     }
    
        
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