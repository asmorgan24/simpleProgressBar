# simpleProgressBar
Macro for PowerPoint to add a progress bar to the bottom or top of the slide deck

## Intro
To view the progress of a Powerpoint presentation, a progress bar can be displayed at the top/bottom of the slide show. I have written code that provides a progress bar featuring two different options:
- A simple progress bar for sectionless Powerpoint slides
- A more advanced progress bar that takes section headers and lengths into account. 

## How to proceed
Once the slideshow is complete, go to **Tools > Macro > Visual Basic Editor**. 

In the new window, select **Insert > Module** and copy the text in ```SimpleProgressBar.bas```: 
Then go to **File > Close > Return to Microsoft PowerPoint**

In the displayed page of Microsoft Powerpoint, go to:
**Tools > Macro > Macros**, then select *AddProcessBar* and press *Execute*

### How remove the progress bar?
**Tools > Macro > Macros**, then select *RemoveProcessBar* and press *Execute*

## Notes
Once you hit run, the macro will add a progress bar to your powerpoint. If slides change order or are added, you will need to once again run the macro. For this reason, I suggest you run the macro once you are (mostly) done with your presentation. Note that hidden slides will be skipped. Also note that I suggest you save your presentation as a "Macro-enabled Powerpoint Presentations" so that the code for the macro is saved and you can return/edit the code later as necessary.  
