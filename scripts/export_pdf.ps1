<#
export_pdf.ps1 —— 用本机 Word 把 .docx 导出成**同名同目录**的 .pdf。

深圳大学《大学物理实验》报告技能专用：每次产出/更新 <实验名称>.docx 之后，
都要导出一份 <实验名称>.pdf（名字与 docx 完全一致，只是扩展名不同）。

用法（用 PowerShell 工具执行，不要从 Bash 里调 powershell）：

    & "<技能目录>\scripts\export_pdf.ps1" -Docx "<工作区>\实验-某某\<实验名称>.docx"

默认输出到 docx 同目录、同名 .pdf；要换位置用 -Out。
结束时打印两行，供核对：

    PDF=<导出后的绝对路径>
    PAGES=<Word 统计的页数>

注意：
* 以只读方式打开 docx，不会改动源文件；
* 不用 Add-Type（会被安全策略拦），只用 New-Object -ComObject；
* 必须装有 Microsoft Word（本技能本来就依赖它核对分页）。
#>
param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Out
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Docx)) {
    throw "docx not found: $Docx"
}
$Docx = (Resolve-Path -LiteralPath $Docx).Path
if (-not $Out) {
    $Out = [System.IO.Path]::ChangeExtension($Docx, ".pdf")
} else {
    $Out = [System.IO.Path]::GetFullPath($Out)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$pages = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $true)   # ReadOnly
    $pages = $doc.ComputeStatistics(2)                  # wdStatisticPages
    $doc.ExportAsFixedFormat($Out, 17, $false)          # 17 = wdExportFormatPDF
    $doc.Close($false)
} finally {
    $word.Quit()
}

"PDF=$Out"
"PAGES=$pages"
