# xlsxPoison

Just a PoC to turn xlsx (regular Excel files) into xlsm (Excel file with macro) and slipping inside a macro (vbaProject.bin). The macro can be modified previously with any tool like [EvilClippy](https://github.com/outflanknl/EvilClippy)

Today most teams that work on common projects in companies tend to share files through network shares or cloud solutions like OneDrive. This fact means that low privileged users have the power to add or edit files on locations that are used by other users, so this situation can be turned into an opportunity for lateral movements or even privilege escalation. We can try to take advantage of this situation converting an existent .xlsx file into a .xlsm, so the user can open it and get "pwned" (it's more likely that the user will trust the "Macro" alert if the file is something he knows to be innofensive)

# How it works?
In the end XLSX/XLSM  are just zip files that follows the Office Open XML file format, so we can do an unzip-edit-rezip to add our macro. 

1. Unzip .xlsx
2. Fix `[Content_Types].xml` and `xl\_rels\workbook.xml.rels`
3. Copy the macro to `xl\`
4. Rezip it with the same name but replacing `xlsx` for `xlsm`
5. Set original file as "hidden"
6. Delete the temporal folder created


*Disclaimer: I tested it with a bunch of random .xlsx files. If you find a .xlsx that get corrupted please ping me at issues*

# Usage

```bash
xlsxPoison.exe file.xlsx vbaProject.bin
```
Example:
```bash
xlsxPoison.exe "C:\Users\avispa.marina\Desktop\Macros\target01.xlsx" "C:\Users\avispa.marina\vbaProject.bin"
```

# Author
Juan Manuel Fernández ([@TheXC3LL](https://twitter.com/TheXC3LL))

