/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function () {
    "use strict";

    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {




            // Use this to check whether the new API is supported in the Word client.
            if (Office.context.requirements.isSetSupported("WordApi", "1.2")) {

                console.log('This code is using Word 2016 or greater.');

                // Setup the event handlers for UI.
                $('#loadSelectedImage').click(loadSelectedImageHandler);
                $('#insertImageAtSelection').click(insertImageHandler);



                initCanvas();

            } else {
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires Word 2016 or greater. Check your version of Word and the requirement set version.');
            }
        });
    };



    /*********************/
    /* Canvas functions */
    /*********************/

    function initCanvas() {

        var canvas = document.getElementById('canvas');

        // Check that canvas is supported.  
        if (canvas.getContext) {

            var ctx = canvas.getContext("2d");

            // Flag to indicate when to draw.
            var isDrawing = false;

            // Get mouse coordinates, and draw while the mouse button is down.
            ctx.canvas.addEventListener('mousemove', function (event) {

                var mouseX = event.clientX - ctx.canvas.offsetLeft;
                var mouseY = event.clientY - ctx.canvas.offsetTop;

                // Check that drawing is enabled. This is enabled when the mouse button
                // is down. It is disabled when the mouse button is up/
                if (isDrawing) {
                    ctx.strokeStyle = "red";
                    ctx.lineTo(mouseX, mouseY);
                    ctx.stroke();
                }

                // Show canvas coordinates.
                var status = document.getElementById('status');
                status.innerHTML = mouseX + " | " + mouseY;
            });

            // Turn on drawing to the canvas. Position the drawing path at the 
            // coordinates where the mouse is down.
            ctx.canvas.addEventListener('mousedown', function (event) {

                // Get the mouse coordinates.
                var mouseX = event.clientX - ctx.canvas.offsetLeft;
                var mouseY = event.clientY - ctx.canvas.offsetTop;

                // Move the canvas drawing point to the mouse location.
                ctx.beginPath();
                ctx.moveTo(mouseX, mouseY);

                // Turn on drawing to the canvas.
                isDrawing = true;
            });

            // Turn off drawing to the canvas.
            ctx.canvas.addEventListener('mouseup', function (event) {
                isDrawing = false;
            });
        }
    }

    function loadImageIntoCanvas(base64EncodedImage, imageWidth, imageHeight) {

        var canvas = document.getElementById('canvas');
        //        canvas.height = imageHeight;
        //        canvas.width = imageWidth;

        // Check that canvas is supported.  
        if (canvas.getContext) {

            var ctx = canvas.getContext("2d");

            var image = new Image();

            image.onload = function () {
                ctx.drawImage(image, 0, 0);
            };

            // BUG: Word provides image data as just the Base64 string. It doesn't describe the
            // image encoding format. This is an issue for HTML canvas (and maybe other APIs) since
            // it doesn't detect the encoding. 
            // DOCS: What file format's will insertInlinePicture accept?
            image.src = "data:image/png;base64," + base64EncodedImage.value;
            //image.src = base64EncodedImage.value;
        }
    }


    /*********************/
    /* Word JS functions */
    /*********************/

    // This assumes that a single image was selected.
    function loadSelectedImageHandler() {
        Word.run(function (context) {

                // Create a proxy object for the range that is assumed to contain an image.
                var imageRange = context.document.getSelection();

                // Create a proxy for the collection of images in the range.
                //                var images = imageRange.inlinePictures;

                // Load the selected range.
                context.load(imageRange, 'inlinePictures');

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(function () {
                        // If there is more than one inline picture, then we need to tell the user to choose a single picture.
                        if (imageRange.inlinePictures.items.length === 1) {

                            // Now we have the image.
                            var imageString = imageRange.inlinePictures.items[0].getBase64ImageSrc();

                            // We need the image height and width so we can scale the canvas.
                            // BUG: image height and width 
                            var imageHeight = imageRange.inlinePictures.items[0].height;
                            var imageWidth = imageRange.inlinePictures.items[0].width;

                            return context.sync().then(function () {
                                loadImageIntoCanvas(imageString, imageHeight, imageWidth);
                            });

                        }
                        // 
                        else {
                            throw "You need to select a single image."
                        }

                    });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function insertImageHandler() {

        var canvas = document.getElementById('canvas');

        var base64ImgString;

        // Check that canvas is supported.  
        if (canvas.getContext) {

            // I wonder how to do encoding detection. This is using the default png encoding.
            var pngDataUrl = canvas.toDataURL(); // data uri scheme

            // Extract encoding format info.
            base64ImgString = pngDataUrl.replace('data:image/png;base64,', '');
            //            base64ImgString = pngDataUrl.slice(pngDataUrl.lastIndexOf('data:image/png;base64,'), pngDataUrl.length);
        }

        Word.run(function (context) {

                // Create a proxy object for the range is the current selection.
                var imageRange = context.document.getSelection();



                // Load the selected range.
                context.load(imageRange, 'text');

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(function () {

                        imageRange.insertInlinePictureFromBase64(base64ImgString, Word.InsertLocation.replace);
                        imageRange.select();
                    })
                    .then(context.sync);
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }


})();