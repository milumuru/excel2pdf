#実行パスを取得
[string]$DirectoryPath = Split-Path $PSCommandPath

#拡張子の指定
$FileExtention = "*.xls*"

#指定拡張子のフルパスを取得
$List = Get-ChildItem -Path $DirectoryPath -Recurse -Include $FileExtention
$FileString = $List | Select-object -Property FullName

#取得したファイル数
$FileCount = $FileString.Length

#初期値
$Count = 0

While ($Count -lt $FileCount ) {

    #1ファイルづつ処理する
    $FullFileName = $FileString[$Count].Fullname

    #フォルダパスとファイル名を取得
    $FolderName =Split-Path $FullFileName -Parent
    $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FullFileName)

    #PDF作成後のパスを作成
    $PDFFiilePath = $FolderName + "\" + $FileName + ".pdf"

    $ExcelApplication = New-Object -ComObject Excel.Application
    $ExcelWorkbook = $ExcelApplication.workbooks.open("$FullFileName")
    $ExcelWorkbook.ExportAsFixedFormat(0,"$PDFFiilePath")
    $ExcelApplication.Workbooks.Close()
    $ExcelApplication.Quit()
 
    #カウントアップする
    $Count++
}