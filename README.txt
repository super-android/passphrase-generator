PASSPHRASE GENERATOR - SETUP GUIDE
===================================

FILES IN THIS FOLDER
--------------------
  PassphraseGen.ps1          - The app itself
  Launch-PassphraseGen.vbs   - Silent launcher (no console window)
  passphrase.ico             - App icon
  eff_large_wordlist.txt     - Drop the EFF wordlist here for max entropy
                               Get it from: https://www.eff.org/files/2016/07/18/eff_large_wordlist.txt


HOW TO LAUNCH (no command line needed)
---------------------------------------
Double-click  Launch-PassphraseGen.vbs  to run the app silently.
The console window will not appear.


CREATING A DESKTOP SHORTCUT (so it appears like a normal app)
--------------------------------------------------------------
1. Right-click Launch-PassphraseGen.vbs -> Send to -> Desktop (create shortcut)
2. On the desktop, right-click the new shortcut -> Properties
3. Click "Change Icon..."
4. Click Browse, navigate to C:\Temp\PasswordGen\passphrase.ico, click Open
5. Click OK -> Apply -> OK
6. Rename the shortcut to "Passphrase Generator"

Now it appears on your desktop with the purple lock icon, launches silently,
and only one instance can run at a time.


PINNING TO TASKBAR OR START MENU
----------------------------------
After creating the desktop shortcut above:
- Taskbar: Right-click the shortcut -> Pin to taskbar
- Start:   Right-click the shortcut -> Pin to Start


SINGLE INSTANCE
---------------
If you try to open a second copy, the existing window will be brought
to the foreground automatically instead.
