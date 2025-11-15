# 《水浒全传》章回分割脚本
# 作者：AI助手
# 功能：使用正则表达式按章回分割水浒全传文本文件

# 设置文件路径
$源文件路径 = "d:\AI\books\水浒传\水浒全传.txt"
$正则文件路径 = "d:\AI\books\水浒传\正则.txt"
$输出目录 = "d:\AI\books\水浒传\章回"

# 读取正则表达式
$正则表达式 = Get-Content $正则文件路径 -Raw

# 创建输出目录（如果不存在）
if (-not (Test-Path $输出目录)) {
    New-Item -ItemType Directory -Path $输出目录 -Force
    Write-Host "创建输出目录: $输出目录"
}

# 读取源文件内容
$文件内容 = Get-Content $源文件路径 -Raw -Encoding UTF8

Write-Host "开始分割《水浒全传》..."
Write-Host "使用正则表达式: $正则表达式"

# 使用正则表达式查找所有章回标题
$章回匹配 = [regex]::Matches($文件内容, $正则表达式)

Write-Host "找到 $($章回匹配.Count) 个章回标题"

# 如果没有找到章回，尝试其他可能的格式
if ($章回匹配.Count -eq 0) {
    Write-Host "未找到章回标题，尝试其他格式..."
    # 尝试其他可能的章回格式
    $备用正则 = "第[一二三四五六七八九十百零\d]+回"
    $章回匹配 = [regex]::Matches($文件内容, $备用正则)
    Write-Host "使用备用正则找到 $($章回匹配.Count) 个章回标题"
}

# 分割文件
for ($i = 0; $i -lt $章回匹配.Count; $i++) {
    $当前章回 = $章回匹配[$i]
    $当前标题 = $当前章回.Value
    $当前位置 = $当前章回.Index
    
    # 确定章回结束位置（下一个章回开始或文件末尾）
    if ($i -lt $章回匹配.Count - 1) {
        $下一个位置 = $章回匹配[$i + 1].Index
        $章回内容 = $文件内容.Substring($当前位置, $下一个位置 - $当前位置).Trim()
    } else {
        $章回内容 = $文件内容.Substring($当前位置).Trim()
    }
    
    # 生成文件名（提取章回号）
    $章回号 = [regex]::Match($当前标题, "第(.{0,5}?)回").Groups[1].Value
    if (-not $章回号) {
        $章回号 = "{0:D3}" -f ($i + 1)
    }
    
    # 确保文件名是3位数字格式
    if ($章回号 -match "^\d+$") {
        $文件名 = "{0:D3}.txt" -f [int]$章回号
    } else {
        $文件名 = "{0:D3}.txt" -f ($i + 1)
    }
    
    $输出文件路径 = Join-Path $输出目录 $文件名
    
    # 写入章回内容
    $章回内容 | Out-File -FilePath $输出文件路径 -Encoding UTF8 -Force
    
    Write-Host "已生成: $文件名 - $当前标题"
}

Write-Host "分割完成！共生成 $($章回匹配.Count) 个章回文件。"
Write-Host "文件保存在: $输出目录"