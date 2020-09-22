
***************************************************************************
clsTranslator - by: Agam Saran
***************************************************************************




Overview
-------------------------------------------------------------
clsTranslator is a class that will allow you to easily, very easily indeed, add Multilingual UI feature to your apps. It fulfils all the things that are helpful in implimenting this feature. Using it is easy too: the syntax of the language files it adopts is very simple, only two lines of code are required to use it and you do not have to make any changes to your app in order to use it (like inserting tags to each control). Being feature-ful and easy, it is fast too. I first developed it for my "CoolWeb" but when it became a feature-ful class, I thought that it should be a separate submission.



Features
-------------------------------------------------------------
1. No changes required to your app.
2. Very easy-to-use.
3. Supports control arrays.
4. Supports "Strings" which are helpful when you do not need to translate a control's Caption or ToolTip but a piece of "String".
5. Supports changing the Captions or ToolTips of Buttons, Panels, Tabs and Column-Headers in respective controls.
6. Chooses the right property for the right control.


Translation File Structure
-------------------------------------------------------------
Comments in translation files can begin with a semi-colon (';'). Strings are added by inserting the keyword "String" followed by a number (which is the ID of the String) and a equal sign ('=') and then by the value of String. Like this:

String19=This is some text

The Captions and ToolTips of controls are given by first telling the Form the control is placed in. This is done by inserting the Form name within '[' and ']'. Like this:

[frmMain]

Afterwards, you have to give the name of the control whose Caption or ToolTip you would like to set and then insert a equal-sign ('=') followed by the Caption of the control. Then put-in a special symbol " | " (without qoutes) and then ToolTip of the control. Like this:

cmdExit=&Exit | Exit using my app and return to Desktop

The class does not set only Caption, it determines the control-type and sets the property accordingly. For Example, if the control is a TextBox or a ComboBox then it sets the Text property instead of Caption. If there is a control array then you can just specify the index in parenthesis i.e '(' and ')' like you do in VB. If there is a control array and you do not specify index then this class sets the property of all controls in the array. The class also supports changing Captions and ToolTips of Buttons of Toolbars, Panels of StatusBars, Tabs of TabStrips and Column-Headers of ListViews. To do this just add the index of the repective object in curly braces before the object name. Like this:

{1}tbrButtons=Bold | Bold the selected text

The above code would change the Caption and ToolTip of first button of "tbrButtons" toolbar.