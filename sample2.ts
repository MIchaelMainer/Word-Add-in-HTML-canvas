
// TODO: find the typescript definition file location and add here.
///<reference path="//appsforoffice.microsoft.com/lib/1.1/hosted/office.d.ts" />
class App {

    constructor() {
        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the new API is supported in the Word client.
                if (Office.context.requirements.isSetSupported("WordApi", "1.2")) {

                    console.log('This code is using Word 2016 or greater.');

                    // Setup the event handlers for UI.
                    $('#loadSelectedImage').click(loadSelectedImageHandler);
                    $('#insertImageAtSelection').click(insertImageHandler);

                    // Scale the size of the canvas so that it scales  
                    // when a user resizes the add-in.
                    window.addEventListener('resize', resizeCanvas, false);

                    // Setup the canvas event listener(s).
                    initCanvas();

                } else {
                    // Just letting you know that this code will not work with your version of Word.
                    console.log('This add-in requires Word 2016 or greater. Check your version of Word and the requirement set version.');
                }
            });
        }
    }
    
    private _calloutEnabled: boolean; // we only want to add callout when an image is loaded.
    private _calloutNumber: number; //set/reset when an image is loaded.
    private _resizeRatio: number; // set when an image has been loaded in to the canvas.
    private _windowWidth: number; // we are setting the canvas width to the window width.
    private _image; // the image added to the canvas.
    
    function initCanvas() {
            
        var canvas: HTMLCanvasElement = document.getElementById('canvas');
        var ctx: CanvasRenderingContext2D = canvas.getContext("2d");
        
        // Add callouts when the user clicks in the canvas.
        ctx.canvas.addEventListener('click', function (event) {

            // Let's make sure that we have an image loaded before
            // we add callouts to the canvas.
            if (_calloutEnabled) {

                // Increment callout number. We will use this later when we stub out
                // descriptions for the callouts by using the Word JS API.
                _calloutNumber++;

                // Get the bounds of the canvas element in relationship to the top-left of the viewport.
                // We will get the coordinates in canvas, not the window.
                var canvasBounds = canvas.getBoundingClientRect();

                // Use the event coordinates, canvas boundaries, and the window width
                // to get the coordinates where the callouts are to be placed.
                var height: number = _windowWidth * _resizeRatio;
                var mouseX: number = (event.clientX - canvasBounds.left) * canvas.width / _windowWidth;
                var mouseY: number = (event.clientY - canvasBounds.top) * canvas.height / height;

                // Draw circle for the callout.
                var radius: number = 12;
                ctx.fillStyle = 'red';
                ctx.beginPath();
                ctx.arc(mouseX, mouseY, radius, 0, Math.PI * 2, true);
                ctx.closePath();
                ctx.fill();

                // Insert the callout number in the circle.
//                var width: number = ctx.measureText(_calloutNumber);
                ctx.font = 'bold 16px calabri ';
                ctx.fillStyle = 'white';
                ctx.textAlign = 'center';
                ctx.fillText(_calloutNumber, mouseX, mouseY + (radius / 3)); 
                // this last argument is approximately correct for placement.
            }
        });
    }

// TODO: continue typescript learning here. 


    // Scale the canvas to fit the add-in window.
    function resizeCanvas() {

        // Canvas must fit width of add-in.
        _windowWidth = window.innerWidth;



        // Set the resize ratio only if it hasn't already been captured, 
        // and only if there is an image loaded in the add-in.
        if (!_resizeRatio && _image)
            _resizeRatio = _image.height / _image.width;

        // Resize the canvas only if there is an image loaded into it.
        if (_image) {
            var height = _windowWidth * _resizeRatio;
            var canvas = document.getElementById('canvas');
            canvas.style.width = _windowWidth + 'px';
            canvas.style.height = height + 'px';
        }
    }

    // Loads the image into the HTML canvas.
    function loadImageIntoCanvas(base64EncodedImage) {

        // Create an image and load it onto the canvas, set the canvas to the image
        // dimensions, and draw it on the canvas.
        _image = new Image();
        _image.onload = function () {

            var canvas = document.getElementById('canvas');
            var ctx = canvas.getContext("2d");

            canvas.height = _image.height;
            canvas.width = _image.width;
            ctx.drawImage(_image, 0, 0);

            // Reset the _calloutNumber when I load a new image.
            _calloutNumber = 0;

            // Enable adding callouts to the canvas.
            _calloutEnabled = true;

        };
        // ASSUMPTION: we are assuming only png files. You will need to determine file type.
        // The Word API will include file format information in a future release. 
        _image.src = "data:image/png;base64," + base64EncodedImage.value;

        // Make the canvas scale to the window.
        resizeCanvas();
    }

    /*********************/
    /* Word JS functions */
    /*********************/

    // Load the the selected image from Word into the add-in. 
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

                            // Queue a command to get the image source. 
                            var imageString = imageRange.inlinePictures.items[0].getBase64ImageSrc();

                            // Synchronize the document state by executing the queued commands, 
                            // and return a promise to indicate task completion.
                            return context.sync().then(function () {
                                loadImageIntoCanvas(imageString);
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

    // Insert the image in the canvas into the Word document.
    // Insert the callout placeholders in to the Word document.  
    function insertImageHandler() {

        var canvas = document.getElementById('canvas');

        // Get the data URL for the image in the canvas. 
        var pngDataUrl = canvas.toDataURL(); // data uri scheme

        // Extract the encoding format information. Word only accepts the base64 content. 
        // ASSUMPTION: that this is a png file.
        var base64ImgString = pngDataUrl.replace('data:image/png;base64,', '');

        Word.run(function (context) {

                // Create a proxy object for the range is the current selection.
                var imageRange = context.document.getSelection();

                // Load the selected range.
                context.load(imageRange, 'text');

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(function () {

                        // Queue a command to insert the image into the document.
                        var lastRange = imageRange.insertInlinePictureFromBase64(base64ImgString, Word.InsertLocation.replace);

                        // Queue a command to navigate the UI to the insert picture.
                        lastRange.select();

                        // Queue an indefinite number of commands to insert paragraphs 
                        // based on the number of callouts added to the image. 
                        if (_calloutNumber > 0) {
                            lastRange = lastRange.insertParagraph('Here are your callout descriptions:', Word.InsertLocation.after);

                            for (var i = 0; i < _calloutNumber; i++) {
                                lastRange = lastRange.insertParagraph((i + 1) + ') [enter callout description].', Word.InsertLocation.after);
                            }
                        }
                    })
                    // Synchronize the document state by executing the queued commands.
                    .then(context.sync);
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

}

var app = new App();