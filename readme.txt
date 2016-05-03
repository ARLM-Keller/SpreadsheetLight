Compiling the source code? Create your own (Visual Studio?) solution/project and add these references:

    DocumentFormat.OpenXml (from the Open XML SDK)
    System
    System.Data
    System.Drawing
    System.Windows.Forms
    System.XML
    System.Xml.Linq
    WindowsBase

Then drag and drop everything in the source code folder into your project. Compile. Have tea.

You can get the latest release versions of SpreadsheetLight at
http://spreadsheetlight.com/

You can send feedback, bug findings and goodwill at the website above.
Or use this email address if you like:
support@spreadsheetlight.com

I actually read emails. :)

SpreadsheetLight uses the MIT License. There's a license.txt somewhere in your download.
Basically you can do whatever you want with the source code and the library,
provided that the author(s) [that's me] isn't/aren't liable for damages.
Read the license for the full thing because this summary doesn't do it justice.

Thanks for using SpreadsheetLight!

With gratitude,
Vincent Tan

-=-=-=-=-=-=-=-=-=-=-

Q) How come it doesn't come with the .sln solution and .csproj project files?

Because SpreadsheetLight is originally a commercial product and I reserve the specific right
to keep the strong name key. And the strong name key is indicated in the project file, which
means I would have give the strong name key too.

Also, I feel those files are extra, in that they detract from you reading the source code.
That's what you want the source code for right?

And the source code files are already arranged in such a manner that you can just drag and drop
them into your own solution/project. If that's difficult, then I'd say reading the source code
and understanding it would be harder.

And also that SpreadsheetLight is designed for the end developer to use. Most of the time, this
developer will just use the DLL and be done with it. He/She doesn't have time to learn the intricacies
of Excel, let alone set up a source code project just to compile the code. There's already a
precompiled DLL there!

And also that I don't assume you're using Visual Studio. Who knows, maybe there'll be Mono support.