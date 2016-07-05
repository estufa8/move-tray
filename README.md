# move-tray
Simple tool to move the Form Designer's bottom tray to the right in Visual Studio.
To use it, show the Move Tray window (View->Other Windows->Move tray). Focus a winform in design mode and click the button on that window.
Sometimes when you open that form after using this, it shows only the tray. Click the button again.

If you share your project with someone else, edit form.resx files and change the value in $this.TrayHeight to a smaller value.

I'm sorry, it would be better to have a toolbar button for this, integrated into VS, but I'm not used to create VS extensions and I developed it fast.
