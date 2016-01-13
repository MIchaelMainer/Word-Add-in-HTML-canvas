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
                $('a.toggler').click(function () {
                    $(this).toggleClass('off');
                    calloutEnabled = !calloutEnabled; // switch modes
                });


                initCanvas();

            } else {
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires Word 2016 or greater. Check your version of Word and the requirement set version.');
            }
        });
    };

    // TODO: This is a must; add add-in commands.

    // Consider the picker FileReader open from local fs

    /*********************/
    /* State */
    /*********************/

    //TODO: This can be improved. 
    var calloutEnabled = false; // we only want to add callout when an image is loaded.
    var calloutNumber = 0;

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

                // Get the bounds of the canvas element. This helps us in case the user has scrolled.
                // We will get the coordinates in canvas, not the window.
                // TODO: refactor this code as it is used three times.
                var canvasBounds = canvas.getBoundingClientRect();

                // Get the mouse coordinates.
                var mouseX = event.clientX - canvasBounds.left;
                var mouseY = event.clientY - canvasBounds.top;

                // Check that drawing is enabled. This is enabled when the mouse button
                // is down. It is disabled when the mouse button is up.
                // We also disable if we re in callout mode.
                if (isDrawing && !calloutEnabled) {
                    ctx.strokeStyle = "red";
                    ctx.lineTo(mouseX, mouseY);
                    ctx.stroke();
                }

                //                // Show canvas coordinates.
                //                var status = document.getElementById('status');
                //                status.innerHTML = Math.round(mouseX) + ", " + Math.round(mouseY);
            });

            // Turn on drawing to the canvas. Position the drawing path at the 
            // coordinates where the mouse is down.
            ctx.canvas.addEventListener('mousedown', function (event) {

                // Get the bounds of the canvas element. This helps us in case the user has scrolled.
                // We will get the coordinates in canvas, not the window.
                var canvasBounds = canvas.getBoundingClientRect();

                // Get the mouse coordinates within the canvas.
                var mouseX = event.clientX - canvasBounds.left;
                var mouseY = event.clientY - canvasBounds.top;

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

            // Add callouts
            ctx.canvas.addEventListener('click', function (event) {

                if (calloutEnabled) {

                    calloutNumber++;

                    // TODO: U(se WordJS to get font info.
                    ctx.font = 'bold 16px calabri ';

                    // Get the bounds of the canvas element. This helps us in case the user has scrolled.
                    // We will get the coordinates in canvas, not the window.
                    var canvasBounds = canvas.getBoundingClientRect();

                    // Get the mouse coordinates within the canvas.
                    var mouseX = event.clientX - canvasBounds.left;
                    var mouseY = event.clientY - canvasBounds.top;


                    // Draw circle for the callout.
                    var radius = 12;
                    ctx.fillStyle = 'red';
                    ctx.beginPath();
                    ctx.arc(mouseX, mouseY, radius, 0, Math.PI * 2, true);
                    ctx.closePath();
                    ctx.fill();

                    // Insert the callout number in the circle.
                    var width = ctx.measureText(calloutNumber);
                    ctx.fillStyle = 'white';
                    ctx.textAlign = 'center';
                    // TODO: figure out how to best position vertical alignment of the text.
                    ctx.fillText(calloutNumber, mouseX, mouseY + (radius / 3));


                    // CONSIDER: Open window with text entry for each callout to give description.
                }
            });

        }
    }

    function loadImageIntoCanvas(base64EncodedImage, imageWidth, imageHeight) {

        // Reset the calloutNumber when I load a new image.
        calloutNumber = 0;

        // Enable adding callout to the canvas.
        calloutEnabled = true;

        $(this).toggleClass('off');

        var canvas = document.getElementById('canvas');

        // Check that canvas is supported.  
        if (canvas.getContext) {

            var ctx = canvas.getContext("2d");

            var image = new Image();

            image.onload = function () {
                canvas.height = image.height;
                canvas.width = image.width;
                ctx.drawImage(image, 0, 0);
            };

            // BUG: Word provides image data as just the Base64 string. It doesn't describe the
            // image encoding format. This is an issue for HTML canvas (and maybe other APIs) since
            // it doesn't detect the encoding. 
            // DOCS: What file format's will insertInlinePicture accept?
            // TODO: detect encoding from the image stream we get from WordJS

            // ASSUMPTION: we are assuming only png files. You will need to determine file type.
            image.src = "data:image/png;base64," + base64EncodedImage.value;

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

                            // DEBUG
                            console.log('Width(x): ' + imageRange.inlinePictures.items[0].width + 'pts\n' +
                                'Height(y): ' + imageRange.inlinePictures.items[0].height + 'pts');

                            var imageHeight = imageRange.inlinePictures.items[0].height;
                            var imageWidth = imageRange.inlinePictures.items[0].width;

                            // DEBUG
                            console.log('Width(x): ' + imageWidth + ' pixels\n' +
                                'Height(y): ' + imageHeight + ' pixels');

                            return context.sync().then(function () {
                                loadImageIntoCanvas(imageString, imageWidth, imageHeight);
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

            // Extract encoding format info. ASSUMPTION: that this is a png file.
            base64ImgString = pngDataUrl.replace('data:image/png;base64,', '');
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

                        var lastRange = imageRange.insertInlinePictureFromBase64(base64ImgString, Word.InsertLocation.replace);

                        if (calloutNumber > 0) {
                            lastRange = lastRange.insertParagraph('Here are your callout descriptions:', Word.InsertLocation.after);

                            for (var i = 0; i < calloutNumber; i++) {
                                lastRange = lastRange.insertParagraph((i + 1) + ') [enter callout description].', Word.InsertLocation.after);
                            }
                        }

                        // This moves the UI view to the last paragraph inserted in the document.
                        lastRange.select();
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