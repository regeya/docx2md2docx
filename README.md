# docx2md2docx

![Screenshot of the thing](/img/screenshot.png)

Just a simple Python and PySide6 wrapper for Pandoc and nvim, to convert a Word coc to Markdown, edit in nvim, and convert back to Word, to have a formatted and styled Word document in Affinity Publisher. 

For this to work, you need:
- WSL and WSLg on Win11 (if using in Windows)
- pandoc
- PySide6
- Neovim
   
It doesn't do much, which includes error checking, and is unlikely to ever be packaged up more nicely than this.  I only made it for the thing in the screenshot: I had Word docs coming my way that were
formatted with styles, but then also had manual formatting, including manual line breaks.  I use this script to use Pandoc to convert the .docx to Markdown, fire up Neovim to edit the Markdown doc, and then
use Pandoc again to convert it back to a .docx.  In theory, the .docx should have everything styled using a stylesheet.  I made this so I could import .docx files into Affinity Publisher, which doesn't have
native Markdown (at least not yet) but does have a .docx importer.  It also has the ability to apply styles from a different Affinity Publisher document.  The one monkey wrench in the works is that import
doesn't always replace styles, but instead makes new styles with similar names.  But as I'm usually using this for one customer and use the same font family for the entire job, manual intervention is
a piece of cake.

I only share this so I don't lose it, and in case it comes in handy for anyone.
