



# Office PowerPoint
  
Module to control Microsoft Office PowerPoint  

*Read this in other languages: [English](Manual_OfficePowerPoint.md), [Português](Manual_OfficePowerPoint.pr.md), [Español](Manual_OfficePowerPoint.es.md)*
  
![banner](imgs/Banner_OfficePowerPoint.png o jpg)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


## Description of the commands

### New presentation
  
Create a new Power Point presentation
|Parameters|Description|example|
| --- | --- | --- |

### Open Presentation
  
Open a Power Point presentation.
|Parameters|Description|example|
| --- | --- | --- |
|File||file.pptx|

### Save presentación
  
Save a opened Power Point presentation
|Parameters|Description|example|
| --- | --- | --- |
|Save file||file.pptx|

### Get slide types
  
Get slide types
|Parameters|Description|example|
| --- | --- | --- |
|Resultt||res |

### Add Slide
  
Add a new slide to the presentation
|Parameters|Description|example|
| --- | --- | --- |
|Type|||

### Write in presentation
  
Write in a Power Point presentation.
|Parameters|Description|example|
| --- | --- | --- |
|Slide index||0|
|Id or element name where will write||Title 1|
|Font size||8|
|Align||Center|
|Bold||True|
|Italic||True|
|Underline||True|
|Write text||Lorem ipsum |

### Close presentation
  
Close the presentation that is running
|Parameters|Description|example|
| --- | --- | --- |

### Add Picture
  
Add an image to the presentation.
|Parameters|Description|example|
| --- | --- | --- |
|Image path||image.jpg|
|Slide index||0|
|Position||2,4.5 |
|Height||5|

### Add textbox
  
Add a textbox to the presentation.
|Parameters|Description|example|
| --- | --- | --- |
|Text||lorem ipsum|
|Position||x,y|
|size||width, height|
|Slide index||0|

### Edit text
  
Modify text of existed slide.
|Parameters|Description|example|
| --- | --- | --- |
|Slide index||0|
|Shape||Title 1|
|Font size||8|
|Align||Center|
|Bold||True|
|Italic||True|
|Underline||True|
|Text||Lorem ipsum|

### Get elements from the slide
  
List the elements in the presentation
|Parameters|Description|example|
| --- | --- | --- |
|Slide index||0|
|Save result||res|

### Edit element properties
  
Edit the properties of an element on a slide
|Parameters|Description|example|
| --- | --- | --- |
|Slide index||0|
|Id or element name||Title 1|
|Position||x,y|
|size||width, height|
|Rotation||45.0|
