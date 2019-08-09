# ConsoleLogger
I've found I often want to save transcripts of my Powershell sessions, which you can do with Start-Transcript/Stop-Transcript, but these often add other bits and pieces and sometimes skip out on bits you want in (like, what you type!), and they don't preserve colour. You can screenshot, but then if you want to copy text out of that...

Enter: ConsoleLogger. I've stolen it from a half-baked function I found on the [Powershell dev blog][devblog], and made it a little more user-friendly. It does console logging, with formatting! E.g.:

![Screenshot demonstrating ConsoleLogger](./Screenie.png)

(but this is just a screenshot; ConsoleLogger actually copies the above as HTML)
(don't mind the Powershell errors. I'm just demonstrating that ConsoleLogger preserves colour.)

It comes with 2 handy cmdlets:

```Powershell
Save-ConsoleAsHtml
```
e.g.
```Powershell
Save-ConsoleAsHtml MyConsoleOutput.html
```

You can then attach that file into an email, or onto Slack/Teams/whatever, [like this](./BriefTranscript.html).

```Powershell
Copy-ConsoleAsHtml
```
e.g.
```Powershell
Copy-ConsoleAsHtml 20
```

This will copy the console HTML to the Clipboard, which you can then paste into Outlook/Word/etc. as formatted text:



You generally can't paste into Notepad/Notepad++ etc because they can't handle rich text -- but if you're pasting into something that loses formatting anyway, why not just use your mouse to copy?

These are also valid:
```Powershell
Copy-ConsoleAsHtml
```
(copies the entire Powershell session from the beginning)

```Powershell
Copy-ConsoleAsHtml 0
```
(doesn't copy anything...)

The functions don't do any error-checking, and aren't particularly user-friendly. That's on my to-do list, and also make them a little easier to use (comments and invocation documentation!, also a `-StartAt <text>` switch, which will start the copy from the first time it finds `<text>`).

To use ConsoleLogger, simply unzip the attached file into your local modules folder (e.g., `C:\Users\<user>\Documents\WindowsPowershell\Modules\`, and either
manually run Import-Module ConsoleLogger when you want to use it, or add Import-Module ConsoleLogger to your Powershell profile (use notepad $Profile to edit) so it's always active.

[devblog]: https://devblogs.microsoft.com/powershell/colorized-capture-of-console-screen-in-html-and-rtf/