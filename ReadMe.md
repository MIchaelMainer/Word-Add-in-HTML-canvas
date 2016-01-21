Summary: Learn how you can use Typescript and the HTML Canvas API to extend the image editing capability in a Word document.
Name: Word-Add-in-Typescript-HTML-Canvas

# Image callouts Word add-in sample: load, edit, and insert images 

This Word add-in sample shows you how to:

1. Create a Word add-in with Typescript.
2. Load images from Word into the add-in.
3. Edit images in the add-in by using the HTML canvas API and insert the images into a Word document.
4. Implement add-in commands that both launch an add-in from the ribbon and run a script from a context menu.
5. Use the Office UI Fabric to create a seamless Word user experience.

TODO: Add GIF that shows the sample running. 

## Prerequisites

To use the Image callouts Word ad-in sample, the following are required.

* [node.js](https://nodejs.org) to serve up the docx files.
* [npm](https://www.npmjs.com/) to install the dependencies.
* Word 2016 16.0.6326.0000 or higher, or any client that supports the Word Javascript API. This sample does a requirement check to see if it is running in a supported host.

> Note: Word for Mac 2016 does not support add-in commands at this time. This sample can run on the Mac without the add-in commands. 

TODO: test on a Mac

## Configure the add-in and Word

1. Unzip and run this [registry key](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/AddInCommandsUndark/EnableAppCmdXLWD.zip) to activate the add-in commands feature.
2. Install the Typescript definitions by running ```tsd install``` in the project's root directory at the command line.
3. Install the project dependencies by running ```npm install``` in the project's root directory. 
4.  



MUST have the reg key noted here : https://github.com/OfficeDev/Office-Add-in-Commands-Samples

TODO: note that add-in commands won't work on the Mac or iPad.
TODO: Create my own map.


tsd install
npm install
gulp copy:libs
gulp


For Mac (note that the add-in commands won't work)
1.	Create a folder called “wef” in Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/
2.	Put the developer manifest in the wef folder (Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/wef)
3.	Open word application on Mac and click on inset->”my add-ins” drop down.
