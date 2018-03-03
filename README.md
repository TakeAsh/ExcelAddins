# Excel Addins

## How to install
1. Create blank Excel Macro-Enabled Workbook (*.xlsm).
1. Open Visual Basic Editor.
1. Import `Module1.bas` under `VBAProject/Modules`.
1. Import `ThisWorkBook.cls` under `VBAProject/Microsoft Excel Objects`.
1. Save this workbook as `*.xlsm` .
1. Open `Property` of `ThisWorkBook`.
1. Change `IsAddin` from `False` to `True`.
1. Save this workbook as `*.xlam` under `%AppData%\Microsoft\AddIns`.
1. Select `File / Options`.
1. Select `Add-ins`.
1. Click `Go` button at `Manage: Excel Add-ins`
1. Check the addin.
