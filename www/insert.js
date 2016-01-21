
(function () {
    Office.initialize = function (reason) { 
        $(document).ready(function () {
                         // Use this to check whether the new API is supported in the Word client.
            if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {
                // $('#imageToInsert').load(function() {
                //     var img = document.getElementById('imageToInsert');
                        
                        
                // });
                    
                    
                    
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
            }
        });
    };

})();

             // The initialize function is run each time the page is loaded.


// function contextMenuInsertImg() { 

//                 // Use this to check whether the new API is supported in the Word client.
//                 if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {
//                     // $('#imageToInsert').load(function() {
//                     //     var img = document.getElementById('imageToInsert');
                        
                        
//                     // });
                    
                    
                    
//                         Word.run(function(context){
//                             // Create a proxy object for the range at the current selection.
//                             var imageRange = context.document.getSelection();
                            
//                             // Load the selected range.
//                             context.load(imageRange, 'text');
                
//                             // Synchronize the document state by executing the queued commands, 
//                             // and return a promise to indicate task completion.
//                             return context.sync()
//                                 .then(function(){
//                                     // Queue a command to insert the image into the document.
//                                     var insertedImage = imageRange.insertText('This is text', Word.InsertLocation.replace);
                                    
//                                     // Queue a command to navigate the UI to the insert picture.
//                                     insertedImage.select();
//                                 })
//                                 // Synchronize the document state by executing the queued commands.
//                                 .then(context.sync);
//                         })
//                         .catch((error) => {
//                             console.log('Error: ' + JSON.stringify(error));
//                             if (error instanceof OfficeExtension.Error) {
//                                 console.log('Debug info: ' + JSON.stringify(error.debugInfo));
//                             }
//                         });
                    
                    
                    
                    
                    
                    
                    
                    
                    
//                     console.log('This was hit');
                    
//                     var img = new Image();

//                     img.onload = function () {
//                         var canvas = document.createElement("canvas");
                        
//                         // Do I need this?
//                         canvas.width =this.width;
//                         canvas.height =this.height;

//                         var ctx = canvas.getContext("2d");
//                         ctx.drawImage(this, 0, 0);

//                         var pngDataUrl = canvas.toDataURL();

//                         var base64ImgString = pngDataUrl.replace('data:image/png;base64,', '');
                        
//                         Word.run(function(context){
//                             // Create a proxy object for the range at the current selection.
//                             var imageRange = context.document.getSelection();
                            
//                             // Load the selected range.
//                             context.load(imageRange, 'text');
                
//                             // Synchronize the document state by executing the queued commands, 
//                             // and return a promise to indicate task completion.
//                             return context.sync()
//                                 .then(function(){
//                                     // Queue a command to insert the image into the document.
//                                     var insertedImage = imageRange.insertInlinePictureFromBase64(base64ImgString, Word.InsertLocation.replace);
                                    
//                                     // Queue a command to navigate the UI to the insert picture.
//                                     insertedImage.select();
//                                 })
//                                 // Synchronize the document state by executing the queued commands.
//                                 .then(context.sync);
//                         })
//                         .catch((error) => {
//                             console.log('Error: ' + JSON.stringify(error));
//                             if (error instanceof OfficeExtension.Error) {
//                                 console.log('Debug info: ' + JSON.stringify(error.debugInfo));
//                             }
//                         });

//                         // img.src = 'media/map.png';
//                         img.src = 'https://127.0.0.1:8085/media/map.png';
//                     }
                        
//                 } else {
//                     // Just letting you know that this code will not work with your version of Word.
//                     console.log('This add-in requires Word 2016 or greater. Check your version of Word and the requirement set version.');
//                 } 
//         };
