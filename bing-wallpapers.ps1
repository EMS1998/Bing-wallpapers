# Bing 壁纸下载脚本
# 获取 Bing 每日壁纸（仅下载一张4K超高清壁纸）
# 来源: https://github.com/timothymctim/Bing-wallpapers
#
# 版权 (c) 2015 Tim van de Kamp
# 许可协议: MIT

Param(
    # 指定获取哪个国家/地区的壁纸，默认自动识别
    [ValidateSet('auto', 'ar-XA', 'bg-BG', 'cs-CZ', 'da-DK', 'de-AT',
    'de-CH', 'de-DE', 'el-GR', 'en-AU', 'en-CA', 'en-GB', 'en-ID',
    'en-IE', 'en-IN', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-US',
    'en-XA', 'en-ZA', 'es-AR', 'es-CL', 'es-ES', 'es-MX', 'es-US',
    'es-XL', 'et-EE', 'fi-FI', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR',
    'he-IL', 'hr-HR', 'hu-HU', 'it-IT', 'ja-JP', 'ko-KR', 'lt-LT',
    'lv-LV', 'nb-NO', 'nl-BE', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT',
    'ro-RO', 'ru-RU', 'sk-SK', 'sl-SL', 'sv-SE', 'th-TH', 'tr-TR',
    'uk-UA', 'zh-CN', 'zh-HK', 'zh-TW')][string]$locale = 'auto',

    # 需要下载的壁纸数量，固定为1
    [int]$files = 1,

    # 图片分辨率参数，当前未使用，统一下载UHD（4K）图片
    [ValidateSet('auto', '800x600', '1024x768', '1280x720', '1280x768',
    '1366x768', '1920x1080', '1920x1200', '720x1280', '768x1024',
    '768x1280', '768x1366', '1080x1920')][string]$resolution = 'auto',

    # 壁纸保存的目标文件夹，默认我的图片\Wallpapers
    [string]$downloadFolder = "$([Environment]::GetFolderPath("MyPictures"))\Wallpapers"
)

# 固定只请求一张壁纸数据
[int]$maxItemCount = 1

# 所有支持的市场代码（和 Param 里的 ValidateSet 一致）
$locales = @('ar-XA', 'bg-BG', 'cs-CZ', 'da-DK', 'de-AT',
    'de-CH', 'de-DE', 'el-GR', 'en-AU', 'en-CA', 'en-GB', 'en-ID',
    'en-IE', 'en-IN', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-US',
    'en-XA', 'en-ZA', 'es-AR', 'es-CL', 'es-ES', 'es-MX', 'es-US',
    'es-XL', 'et-EE', 'fi-FI', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR',
    'he-IL', 'hr-HR', 'hu-HU', 'it-IT', 'ja-JP', 'ko-KR', 'lt-LT',
    'lv-LV', 'nb-NO', 'nl-BE', 'nl-NL', 'pl-PL', 'pt-BR', 'pt-PT',
    'ro-RO', 'ru-RU', 'sk-SK', 'sl-SL', 'sv-SE', 'th-TH', 'tr-TR',
    'uk-UA', 'zh-CN', 'zh-HK', 'zh-TW')

# 随机选一个国家/地区代码
if ($locale -eq 'auto') {
    $locale = Get-Random -InputObject $locales
    Write-Host "随机选择地区: $locale"
}

$market = "&mkt=$locale"

# Bing 域名和API请求URL
[string]$hostname = "https://www.bing.com"
[string]$uri = "$hostname/HPImageArchive.aspx?format=xml&idx=0&n=$maxItemCount$market"

# 如果下载文件夹不存在，则创建
if (!(Test-Path $downloadFolder)) {
    New-Item -ItemType Directory $downloadFolder | Out-Null
}

# 请求Bing壁纸的XML信息
$request = Invoke-WebRequest -Uri $uri
[xml]$content = $request.Content

# 新建列表存放图片信息
$items = New-Object System.Collections.ArrayList

# 遍历XML中的每个图片节点
foreach ($xmlImage in $content.images.image) {
    # 解析图片发布日期
    [datetime]$imageDate = [datetime]::ParseExact($xmlImage.startdate, 'yyyyMMdd', $null)

    # 构造4K超高清图片地址，使用_UHD.jpg后缀
    [string]$imageUrl = "$hostname$($xmlImage.urlBase)_UHD.jpg"

    # 创建自定义对象存储日期和URL
    $item = New-Object System.Object
    $item | Add-Member -Type NoteProperty -Name date -Value $imageDate
    $item | Add-Member -Type NoteProperty -Name url -Value $imageUrl

    # 添加到列表
    $null = $items.Add($item)
}

# 保留最新的 $files 张图片（此处固定为1）
if (!($files -eq 0) -and ($items.Count -gt $files)) {
    # 按日期升序排序，删除多余的旧图片
    $items = $items | Sort-Object date
    while ($items.Count -gt $files) {
        $null, $items = $items
    }
}

Write-Host "开始下载壁纸..."
$client = New-Object System.Net.WebClient

# 遍历待下载图片列表
foreach ($item in $items) {
    # 根据日期生成文件名
    $baseName = $item.date.ToString("yyyy-MM-dd")
    $destination = Join-Path $downloadFolder "$baseName.jpg"
    $url = $item.url

    # 如果文件不存在，才下载
    if (!(Test-Path $destination)) {
        Write-Host "下载中：$url -> $destination"
        $client.DownloadFile($url, $destination)
    } else {
        Write-Host "文件已存在，跳过下载：$destination"
    }
}

# 清理旧壁纸，最多保留 $files 张（此处为1张）
if ($files -gt 0) {
    Write-Host "开始清理旧壁纸..."
    $i = 1
    Get-ChildItem -Filter "????-??-??.jpg" -Path $downloadFolder | Sort-Object -Property Name -Descending | ForEach-Object {
        if ($i -gt $files) {
            Write-Host "删除旧文件：$($_.FullName)"
            Remove-Item $_.FullName
        }
        $i++
    }
}
# 获取最新下载的图片文件（按名称排序）
$latestWallpaper = Get-ChildItem -Path $downloadFolder -Filter "????-??-??.jpg" | Sort-Object -Property Name -Descending | Select-Object -First 1

# 如果找到了文件，就设置为壁纸
if ($latestWallpaper) {
    $path = $latestWallpaper.FullName

    # 设置壁纸（调用Windows API）
    Add-Type @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

    # 设置壁纸（SPI_SETDESKWALLPAPER = 20, 更新用户设置 = 1, 写入ini = 2）
    [Wallpaper]::SystemParametersInfo(20, 0, $path, 3)

    Write-Host "壁纸已设置为：$path"
} else {
    Write-Host "未找到要设置的壁纸文件"
}
