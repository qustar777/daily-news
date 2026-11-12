# -*- coding: utf-8 -*-
<#
每日新闻采集脚本 (PowerShell)
信源：中文第一梯队官方媒体
处理：分类、去重、重要性打分、提取式摘要
输出：data/YYYY-MM-DD.json + data/index.json
#>

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# 1. 信源配置
# ============================================================
$SOURCES = @(
    # ---------- RSS 源（结构稳定，直接解析） ----------
    @{ name = "中国新闻网"; type = "rss"; url = "https://www.chinanews.com.cn/rss/scroll-news.xml"; weight = 4; category = ""; maxAge = 48 },

    # ---------- 网页列表页源（RSS 已废弃，改抓 HTML 列表） ----------
    # 央视新闻（频道页为JS渲染，改用首页，按关键词再分类）
    @{ name = "央视新闻"; type = "html"; url = "https://news.cctv.com/"; weight = 5; category = ""; linkPattern = '/\d{4}/\d{2}/\d{2}/ARTI'; maxAge = 48 },

    # 新华社（首页综合）
    @{ name = "新华社"; type = "html"; url = "https://www.news.cn/"; weight = 5; category = ""; linkPattern = '/\d{8}/'; maxAge = 48 },

    # 人民日报（国际频道）
    @{ name = "人民日报"; type = "html"; url = "http://world.people.com.cn/"; weight = 5; category = "国际"; linkPattern = '/n1/\d{4}/'; maxAge = 48 },

    # ---------- 政策类信源（更新频率低，窗口放宽到7天） ----------
    @{ name = "中国人民银行"; type = "html"; url = "http://www.pbc.gov.cn/goutongjiaoliu/113456/113469/index.html"; weight = 5; category = "经济"; linkPattern = '/\d{8,}/index\.html'; maxAge = 168 },

    @{ name = "国家统计局"; type = "html"; url = "https://www.stats.gov.cn/sj/zxfb/"; weight = 5; category = "经济"; linkPattern = 't\d{8}_'; maxAge = 168 }
)

# ============================================================
# 2. 板块关键词
# ============================================================
$CATEGORY_KEYWORDS = @{
    "经济" = @("GDP","CPI","PPI","利率","央行","降准","降息","加息","汇率","股市","沪指","深成指","创业板","油价","黄金","通胀","财政","赤字","国债","税收","贸易","关税","出口","进口","就业","失业","工资","消费","零售","投资","房地产","房价","货币政策","财政政策","LPR","MLF","逆回购","外汇","人民币","美元","欧元","A股","进博会","博览会","自贸","供应链","物流","工业利润","规模以上","制造业","企业利润","金融市场","贷款","融资","互换协议")
    "政治" = @("外交","峰会","元首","总统","主席","总理","部长","大使","选举","投票","政党","国会","议会","法案","宪法","制裁","条约","协定","联合国","外交部","使馆","领事","主权","领土","统一","独立","抗议","示威","革命","政变","讲话","会见","访问")
    "国际" = @("美国","俄罗斯","欧盟","欧洲","日本","韩国","朝鲜","印度","中东","以色列","巴勒斯坦","乌克兰","北约","G7","G20","联合国","WTO","IMF","世界银行","东盟","金砖","会晤","声明","决议","白宫","克里姆林宫")
    "科技" = @("芯片","半导体","光刻机","人工智能","AI","大模型","算法","航天","火箭","卫星","发射","空间站","探月","火星","5G","6G","量子","区块链","新能源","电池","光伏","专利","研发","实验室","科学家","突破","科技","技术","互联网","数字")
    "军事" = @("军演","演习","导弹","国防","航母","舰艇","战机","坦克","冲突","战争","停火","军援","武器","弹药","部队","解放军","美军","俄军","北约","防务","战备","海军","空军","陆军")
    "社会" = @("灾害","地震","洪水","台风","疫情","传染病","疫苗","事故","救援","应急","安全","交通","铁路","航空","教育","医疗","环保","民生","城市","农村","人口")
}

# 重磅关键词（命中加分）
$IMPORTANT_KEYWORDS = @("利率决议","加息","降息","降准","战争","制裁","峰会","元首会晤","IPO","退市","破产","地震","疫情","发射","军演","导弹","停火","和谈","选举","政变","冲突","签署","协定")

$MAX_PER_CATEGORY = 10

# ============================================================
# 3. 工具函数
# ============================================================
function Get-TextSimilarity {
    param([string]$a, [string]$b)
    if ([string]::IsNullOrEmpty($a) -or [string]::IsNullOrEmpty($b)) { return 0 }
    # 归一化：去掉标点和空白
    $na = $a -replace '[^\u4e00-\u9fa5a-zA-Z0-9]', ''
    $nb = $b -replace '[^\u4e00-\u9fa5a-zA-Z0-9]', ''
    if ($na.Length -eq 0 -or $nb.Length -eq 0) { return 0 }
    # 长度差异过大直接判不相似
    $ratio = [math]::Min($na.Length, $nb.Length) / [math]::Max($na.Length, $nb.Length)
    if ($ratio -lt 0.55) { return 0 }

    # 生成二元组（bigram）集合
    $gramsA = @{}
    for ($i = 0; $i -lt $na.Length - 1; $i++) { $gramsA[$na.Substring($i, 2)] = $true }
    $gramsB = @{}
    for ($i = 0; $i -lt $nb.Length - 1; $i++) { $gramsB[$nb.Substring($i, 2)] = $true }
    if ($gramsA.Count -eq 0 -or $gramsB.Count -eq 0) { return 0 }

    $inter = 0
    foreach ($g in $gramsA.Keys) { if ($gramsB.ContainsKey($g)) { $inter++ } }
    $union = $gramsA.Count + $gramsB.Count - $inter
    if ($union -eq 0) { return 0 }
    return [math]::Round($inter / $union, 3)
}

function Get-TitleHash {
    param([string]$title)
    # 归一化后的完整标题哈希（只用于识别完全相同的标题）
    $norm = $title -replace '[^\u4e00-\u9fa5a-zA-Z0-9]', ''
    if ($norm.Length -eq 0) { $norm = $title }
    $bytes = [Text.Encoding]::UTF8.GetBytes($norm)
    return (Get-FileHash -InputStream ([IO.MemoryStream]::new($bytes)) -Algorithm MD5).Hash
}

function Get-Category {
    param([string]$title, [string]$content)
    $text = $title + " " + ($content.Substring(0, [Math]::Min(200, $content.Length)))
    $bestCat = "其他"
    $bestScore = 0
    foreach ($cat in $CATEGORY_KEYWORDS.Keys) {
        $score = 0
        foreach ($kw in $CATEGORY_KEYWORDS[$cat]) {
            if ($text.Contains($kw)) { $score++ }
        }
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestCat = $cat
        }
    }
    return @{ category = $bestCat; score = $bestScore }
}

function Get-ImportanceScore {
    param($item)
    $score = 0
    $score += $item.source_weight * 2
    $title = $item.title
    $content = if ($item.content) { $item.content.Substring(0, [Math]::Min(300, $item.content.Length)) } else { "" }
    foreach ($kw in $IMPORTANT_KEYWORDS) {
        if ($title.Contains($kw) -or $content.Contains($kw)) { $score += 3 }
    }
    $score += $item.category_score
    $contentLen = if ($content) { $content.Length } else { 0 }
    if ($contentLen -gt 50 -and $contentLen -lt 2000) { $score += 1 }
    return $score
}

function Get-Summary {
    param([string]$title, [string]$content)
    if ([string]::IsNullOrWhiteSpace($content)) { return $title }
    $clean = $content -replace '\s+', ''
    $sentences = $clean -split '[。！？\n]' | Where-Object { $_.Length -gt 10 }
    if ($sentences.Count -eq 0) {
        return if ($clean.Length -gt 120) { $clean.Substring(0, 120) + "..." } else { $clean }
    }
    $first = $sentences[0]
    $skipPatterns = @("本报讯","记者","新华社","央视网","人民网","据","中新网","环球网")
    foreach ($p in $skipPatterns) {
        if ($first.StartsWith($p)) {
            if ($sentences.Count -gt 1) { $first = $sentences[1] }
            break
        }
    }
    if ($first.Length -gt 120) {
        return $first.Substring(0, 120) + "..."
    }
    return $first
}

function Strip-Html {
    param([string]$html)
    if ([string]::IsNullOrEmpty($html)) { return "" }
    $text = $html -replace '<[^>]+>', ''
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return $text.Trim()
}

function Test-IsRecent {
    param([string]$pubTimeStr, [int]$hours = 48)
    if ([string]::IsNullOrWhiteSpace($pubTimeStr)) { return $true }  # 无时间戳暂保留，后面用URL日期兜底
    $now = Get-Date
    $parsed = $null
    # 尝试多种格式
    $formats = @(
        "ddd, d MMM yyyy HH:mm:ss zzz",
        "ddd, d MMM yyyy HH:mm:ss GMT",
        "yyyy-MM-dd HH:mm:ss",        "yyyy-MM-dd HH:mm",
        "yyyy-MM-dd",
        "yyyy/MM/dd HH:mm",
        "yyyy/MM/dd"
    )
    foreach ($fmt in $formats) {
        try {
            $parsed = [DateTime]::ParseExact($pubTimeStr.Trim(), $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            break
        } catch {}
    }
    if (-not $parsed) {
        # 用正则提取日期
        $m = [regex]::Match($pubTimeStr, '(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
        if ($m.Success) {
            try {
                $parsed = Get-Date -Year $m.Groups[1].Value -Month $m.Groups[2].Value -Day $m.Groups[3].Value
            } catch {}
        }
    }
    if (-not $parsed) { return $true }  # 解析不出时间，暂保留
    $delta = $now - $parsed
    return ($delta.TotalHours -le $hours)
}

# 把 pub_time 统一格式化成 "yyyy-MM-dd"，便于前端按发布日期精确过滤
function Format-PubDate {
    param([string]$pubTimeStr)
    if ([string]::IsNullOrWhiteSpace($pubTimeStr)) { return $null }
    $formats = @(
        "ddd, d MMM yyyy HH:mm:ss zzz",
        "ddd, d MMM yyyy HH:mm:ss GMT",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-dd HH:mm",
        "yyyy-MM-dd",
        "yyyy/MM/dd HH:mm",
        "yyyy/MM/dd"
    )
    foreach ($fmt in $formats) {
        try {
            $dt = [DateTime]::ParseExact($pubTimeStr.Trim(), $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            return $dt.ToString('yyyy-MM-dd')
        } catch {}
    }
    # 正则兜底
    $m = [regex]::Match($pubTimeStr, '(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
    if ($m.Success) {
        try {
            $dt = Get-Date -Year $m.Groups[1].Value -Month $m.Groups[2].Value -Day $m.Groups[3].Value
            return $dt.ToString('yyyy-MM-dd')
        } catch {}
    }
    return $null
}

function Get-DateFromUrl {
    param([string]$url)
    if ([string]::IsNullOrWhiteSpace($url)) { return $null }

    # 格式1：/2026/11/01/  （央视）
    $m = [regex]::Match($url, '/(\d{4})/(\d{2})/(\d{2})/')
    if ($m.Success) {
        try { return Get-Date -Year ([int]$m.Groups[1].Value) -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[3].Value) } catch {}
    }

    # 格式2：/n1/2026/1101/  （人民日报，年+月日）
    $m = [regex]::Match($url, '/(\d{4})/(\d{4})/')
    if ($m.Success) {
        try {
            $y = [int]$m.Groups[1].Value; $md = $m.Groups[2].Value
            return Get-Date -Year $y -Month ([int]$md.Substring(0,2)) -Day ([int]$md.Substring(2,2))
        } catch {}
    }

    # 格式3：8位连续数字 20261101（新华社 /20261101/、统计局 t20261031_、央行 2026103018...）
    $m = [regex]::Match($url, '(\d{4})(\d{2})(\d{2})')
    if ($m.Success) {
        try {
            $y = [int]$m.Groups[1].Value; $mo = [int]$m.Groups[2].Value; $d = [int]$m.Groups[3].Value
            if ($mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
                return Get-Date -Year $y -Month $mo -Day $d
            }
        } catch {}
    }

    # 格式4：/2026-11-01/
    $m = [regex]::Match($url, '(\d{4})-(\d{2})-(\d{2})')
    if ($m.Success) {
        try { return Get-Date -Year ([int]$m.Groups[1].Value) -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[3].Value) } catch {}
    }

    return $null
}

# 兼容 PowerShell 5.1 与 7.x：-UseBasicParsing 在 7.x 已移除
function Invoke-WebSafe {
    param([string]$url, [int]$timeout = 15)
    $p = @{ Uri = $url; Headers = $headers; TimeoutSec = $timeout }
    if ($PSVersionTable.PSVersion.Major -lt 6) { $p['UseBasicParsing'] = $true }
    return Invoke-WebRequest @p
}

function Fetch-ArticleContent {
    param([string]$url)
    if ([string]::IsNullOrWhiteSpace($url)) { return "" }
    try {
        $resp = Invoke-WebSafe -url $url -timeout 15
        $html = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        if ($html.Length -lt 500) { return "" }

        # 第零步：剥离 <style> 和 <script> 块，避免 CSS/JS 混入正文
        $cleanHtml = [regex]::Replace($html, '<script[\s\S]*?</script>', '', 'IgnoreCase')
        $cleanHtml = [regex]::Replace($cleanHtml, '<style[\s\S]*?</style>', '', 'IgnoreCase')

        # 第一步：尝试定位正文容器
        $containerPatterns = @(
            '<div[^>]*class="[^"]*(?:article-content|content-detail|detail-content|article_body|content_main|articletext|main-content|news-content|content_text|text_con)["\s]',
            '<div[^>]*id="[^"]*(?:articleContent|artibody|contentMain|articleBody|article-content)["\s]',
            '<div[^>]*class="content"[^>]*>'
        )
        $containerHtml = $null
        foreach ($pat in $containerPatterns) {
            $m = [regex]::Match($cleanHtml, $pat, 'IgnoreCase')
            if ($m.Success) {
                $start = $m.Index
                $openTagEnd = $cleanHtml.IndexOf('>', $start)
                if ($openTagEnd -ge 0) {
                    $cand = Extract-BalancedDiv $cleanHtml $openTagEnd
                    # 验证容器内确实有足够的段落文本，避免匹配到导航/广告容器
                    if ($cand -and $cand.Length -gt 200) {
                        $pCount = ([regex]::Matches($cand, '<p>', 'IgnoreCase')).Count
                        if ($pCount -ge 2) { $containerHtml = $cand; break }
                    }
                }
            }
        }

        # 第二步：从容器内提取 <p>，否则退而求其次抓全页 <p>
        $sourceHtml = if ($containerHtml) { $containerHtml } else { $cleanHtml }
        $pMatches = [regex]::Matches($sourceHtml, '<p[^>]*>([\s\S]*?)</p>', 'IgnoreCase')
        $paras = @()
        foreach ($p in $pMatches) {
            $txt = Strip-Html $p.Groups[1].Value
            if ($txt.Length -lt 15) { continue }
            # 过滤疑似相关新闻标题：短且不以标点结尾
            if ($txt.Length -le 40 -and $txt -notmatch '[。！？.!?；;，,]$') { continue }
            # 过滤明显的噪声行
            if ($txt -match '^(点击|查看|阅读|更多|展开|收起|责任编辑|责编|编辑|来源|作者|版权|Copyright|上一篇|下一篇|分享到|推荐阅读|相关新闻|热门新闻)') { continue }
            $paras += $txt
        }

        $bestText = if ($paras.Count -gt 0) { ($paras -join "`n`n") } else { "" }

        # 兜底：去掉所有HTML标签后取正文
        if ($bestText.Length -lt 200) {
            $clean = Strip-Html $cleanHtml
            $bestText = $clean
        }

        # 第三步：边界截断
        $cutMarkers = @('责任编辑', '责编', '编辑：', '编辑:', '相关新闻', '推荐阅读', '热门新闻', '更多新闻',
                         '分享到', '上一篇', '下一篇', 'Copyright', '版权所有', '版权声明', '点击查看',
                         '相关链接', '延伸阅读', '点击进入', '分享至', '来源：', '作者：')
        foreach ($marker in $cutMarkers) {
            $idx = $bestText.IndexOf($marker)
            if ($idx -gt 100) {
                $bestText = $bestText.Substring(0, $idx)
                break
            }
        }

        # 限制长度
        if ($bestText.Length -gt 3000) {
            $bestText = $bestText.Substring(0, 3000) + "..."
        }
        return $bestText.Trim()
    } catch {
        return ""
    }
}

function Extract-BalancedDiv {
    param([string]$html, [int]$openTagEnd)
    # 从 <div ...> 的 '>' 位置开始，提取匹配的 <div>...</div> 块
    $depth = 1
    $pos = $openTagEnd + 1
    while ($pos -lt $html.Length -and $depth -gt 0) {
        $nextOpen = $html.IndexOf('<div', $pos, 'IgnoreCase')
        $nextClose = $html.IndexOf('</div>', $pos, 'IgnoreCase')
        if ($nextClose -lt 0) { break }
        if ($nextOpen -ge 0 -and $nextOpen -lt $nextClose) {
            $depth++
            $pos = $nextOpen + 4
        } else {
            $depth--
            if ($depth -eq 0) {
                return $html.Substring($openTagEnd + 1, $nextClose - $openTagEnd - 1)
            }
            $pos = $nextClose + 6
        }
    }
    return $null
}

function Fetch-HtmlList {
    param($source)
    $items = @()
    try {
        $resp = Invoke-WebSafe -url $source.url -timeout 15
        $html = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())

        # 提取所有 <a href="...">标题</a>
        $matches = [regex]::Matches($html, '<a[^>]+href="([^"]+)"[^>]*>([^<]{8,90})</a>')
        $seen = @{}
        $count = 0

        foreach ($m in $matches) {
            if ($count -ge 30) { break }
            $href = $m.Groups[1].Value
            $text = (Strip-Html $m.Groups[2].Value).Trim()

            # 过滤：标题必须像正文标题，排除导航/参数模板
            if ($text.Length -lt 10) { continue }
            if ($text -match '\+data|clickUrl|datalist|^[A-Za-z0-9+.,]+$') { continue }
            if ($href -match 'javascript:|#') { continue }

            # 按信源指定的链接模式过滤
            if ($source.linkPattern -and $href -notmatch $source.linkPattern) { continue }

            # 补全相对URL
            $fullUrl = $href
            if ($href.StartsWith("./")) {
                $fullUrl = $source.url.TrimEnd('/') + '/' + $href.Substring(2)
            } elseif ($href.StartsWith("/")) {
                $base = [System.Uri]$source.url
                $fullUrl = $base.Scheme + "://" + $base.Host + $href
            } elseif ($href -notmatch '^http') {
                $fullUrl = $source.url.TrimEnd('/') + '/' + $href
            }

            if ($seen.ContainsKey($fullUrl)) { continue }
            $seen[$fullUrl] = $true

            $urlDate = Get-DateFromUrl $fullUrl
            $pubTime = if ($urlDate) { $urlDate.ToString("yyyy-MM-dd") } else { "" }

            $items += [PSCustomObject]@{
                title = $text
                link = $fullUrl
                content = ""
                pub_time = $pubTime
                source = $source.name
                source_weight = $source.weight
                hint_category = $source.category
                max_age = $source.maxAge
            }
            $count++
        }
    } catch {
        Write-Host "[列表页错误] $($source.name) $($source.url): $_"
    }
    return $items
}

# ============================================================
# 4. 采集函数
# ============================================================
$headers = @{
    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    "Accept-Language" = "zh-CN,zh;q=0.9,en;q=0.8"
}

function Get-XmlText {
    param($node)
    if ($null -eq $node) { return "" }
    if ($node -is [string]) { return $node }
    if ($node.InnerText) { return $node.InnerText }
    if ($node.'#text') { return $node.'#text' }
    return $node.ToString()
}

function Fetch-Rss {
    param($source)
    $items = @()
    try {
        $resp = Invoke-WebSafe -url $source.url -timeout 15
        # 强制用 UTF-8 解码（很多中文 RSS 的 HTTP 头没声明编码，PS 默认按 Latin-1 解码会乱码）
        $content = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        [xml]$xml = $content
        $entries = $xml.rss.channel.item
        if (-not $entries) { $entries = $xml.feed.entry }
        if (-not $entries) { return $items }

        $count = 0
        foreach ($entry in $entries) {
            if ($count -ge 30) { break }
            $title = ""
            $link = ""
            $pubTime = ""
            $summary = ""

            $title = Get-XmlText $entry.title
            if ($entry.link) {
                $link = Get-XmlText $entry.link
                if (-not $link -and $entry.link.href) { $link = $entry.link.href }
            }
            if ($entry.pubDate) { $pubTime = Get-XmlText $entry.pubDate }
            elseif ($entry.updated) { $pubTime = Get-XmlText $entry.updated }
            elseif ($entry.published) { $pubTime = Get-XmlText $entry.published }

            if ($entry.description) {
                $summary = Strip-Html (Get-XmlText $entry.description)
            } elseif ($entry.summary) {
                $summary = Strip-Html (Get-XmlText $entry.summary)
            } elseif ($entry.content) {
                $summary = Strip-Html (Get-XmlText $entry.content)
            }

            if ($title) {
                $items += [PSCustomObject]@{
                    title = $title
                    link = $link
                    content = $summary
                    pub_time = $pubTime
                    source = $source.name
                    source_weight = $source.weight
                    hint_category = $source.category
                    max_age = $source.maxAge
                }
                $count++
            }
        }
    } catch {
        Write-Host "[RSS错误] $($source.name) $($source.url): $_"
    }
    return $items
}

# ============================================================
# 5. 主流程
# ============================================================
Write-Host "============================================================"
Write-Host "每日新闻采集启动 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "============================================================"

$allNews = @()

# 采集所有信源
foreach ($source in $SOURCES) {
    Write-Host "[采集] $($source.name) - $($source.url)"
    if ($source.type -eq "html") {
        $items = Fetch-HtmlList $source
    } else {
        $items = Fetch-Rss $source
    }
    Write-Host "  -> 获取 $($items.Count) 条"
    $allNews += $items
    Start-Sleep -Milliseconds 300
}

Write-Host "`n[采集完成] 共获取 $($allNews.Count) 条原始新闻"

# 处理每条新闻
$processed = @()
foreach ($item in $allNews) {
    if (-not $item.title) { continue }

    # 时效性过滤：媒体类48小时，政策类信源放宽到7天
    $hours = if ($item.max_age) { [int]$item.max_age } else { 48 }
    $isRecent = $false
    if (-not [string]::IsNullOrWhiteSpace($item.pub_time)) {
        $isRecent = Test-IsRecent $item.pub_time $hours
    } else {
        $urlDate = Get-DateFromUrl $item.link
        if ($urlDate) {
            $isRecent = ((Get-Date) - $urlDate).TotalHours -le $hours
        } else {
            $isRecent = $true  # 实在没有日期，暂保留
        }
    }
    if (-not $isRecent) { continue }

    # 确保每条新闻都有 pub_time：空时从 URL 兜底
    if ([string]::IsNullOrWhiteSpace($item.pub_time)) {
        $urlDate = Get-DateFromUrl $item.link
        if ($urlDate) {
            $item.pub_time = $urlDate.ToString('yyyy-MM-dd')
        }
    }

    # 统一 pub_time 格式为 "yyyy-MM-dd"（兼容各种信源的不同格式）
    $normalized = Format-PubDate $item.pub_time
    if ($normalized) { $item.pub_time = $normalized }

    # 抓取文章正文（仅当没有正文时才抓，避免重复）
    if ([string]::IsNullOrWhiteSpace($item.content) -or $item.content.Length -lt 100) {
        $articleContent = Fetch-ArticleContent $item.link
        if ($articleContent) { $item.content = $articleContent }
        Start-Sleep -Milliseconds 200
    }

    # 分类
    $catResult = Get-Category $item.title $item.content
    $category = $catResult.category
    $catScore = $catResult.score

    # 如果信源有板块提示且关键词未命中，用提示的板块
    if ($catScore -eq 0 -and $item.hint_category) {
        $category = $item.hint_category
    }

    # 摘要（太短则用标题代替）
    $summary = Get-Summary $item.title $item.content
    if ($summary.Length -lt 5) { $summary = $item.title }

    $processedItem = [PSCustomObject]@{
        title = $item.title
        summary = $summary
        content = $item.content
        link = $item.link
        source = $item.source
        sources = @($item.source)
        pub_time = $item.pub_time
        category = $category
        category_score = $catScore
        source_weight = $item.source_weight
    }
    $processedItem | Add-Member -NotePropertyName importance_score -NotePropertyValue (Get-ImportanceScore $processedItem)
    $processed += $processedItem
}

# 去重
$seenHashes = @{}
foreach ($item in $processed) {
    $h = Get-TitleHash $item.title
    if ($seenHashes.ContainsKey($h)) {
        $existing = $seenHashes[$h]
        if ($existing.sources -notcontains $item.source) {
            $existing.sources += $item.source
        }
        if ($item.summary.Length -gt $existing.summary.Length) {
            $existing.summary = $item.summary
        }
    } else {
        $seenHashes[$h] = $item
    }
}

# 标题相似度去重
$unique = @()
foreach ($item in $seenHashes.Values) {
    $isDup = $false
    foreach ($u in $unique) {
        $sim = Get-TextSimilarity $item.title $u.title
        if ($sim -ge 0.6) {
            foreach ($s in $item.sources) {
                if ($u.sources -notcontains $s) { $u.sources += $s }
            }
            if ($item.importance_score -gt $u.importance_score) {
                $u.title = $item.title
                $u.summary = $item.summary
                $u.link = $item.link
                $u.importance_score = $item.importance_score
            }
            $isDup = $true
            break
        }
    }
    if (-not $isDup) { $unique += $item }
}

# 按板块分组
$categories = @{}
foreach ($item in $unique) {
    if (-not $categories.ContainsKey($item.category)) {
        $categories[$item.category] = @()
    }
    $categories[$item.category] += $item
}

# “其他”板块只保留有分量的条目：需命中关键词，或来自第一梯队权威信源
if ($categories.ContainsKey("其他")) {
    $categories["其他"] = $categories["其他"] | Where-Object {
        $_.importance_score -ge 11 -or $_.source_weight -ge 5
    }
    if ($categories["其他"].Count -eq 0) { $categories.Remove("其他") }
    elseif ($categories["其他"].Count -gt 3) {
        $categories["其他"] = $categories["其他"] | Sort-Object -Property importance_score -Descending | Select-Object -First 3
    }
}

# 每个板块按重要性排序，并限制每信源最多4条，避免单一信源刷屏
$outputCategories = @{}
$totalNews = 0
foreach ($cat in $categories.Keys) {
    $perSource = @{}
    $picked = @()
    $sortedAll = $categories[$cat] | Sort-Object -Property importance_score -Descending
    foreach ($item in $sortedAll) {
        if ($picked.Count -ge $MAX_PER_CATEGORY) { break }
        $src = $item.source
        $cur = if ($perSource.ContainsKey($src)) { $perSource[$src] } else { 0 }
        if ($cur -ge 4) { continue }
        $perSource[$src] = $cur + 1
        $picked += $item
    }
    # 若因限额导致不足，用剩余条目补齐
    if ($picked.Count -lt $MAX_PER_CATEGORY) {
        foreach ($item in $sortedAll) {
            if ($picked.Count -ge $MAX_PER_CATEGORY) { break }
            if ($picked -notcontains $item) { $picked += $item }
        }
    }
    $cleaned = @()
    foreach ($item in $picked) {
        $cleaned += [PSCustomObject]@{
            title = $item.title
            summary = $item.summary
            content = $item.content
            link = $item.link
            source = $item.source
            sources = $item.sources
            pub_time = $item.pub_time
        }
    }
    $outputCategories[$cat] = $cleaned
    $totalNews += $cleaned.Count
}

# 输出
$today = Get-Date -Format 'yyyy-MM-dd'
$dataDir = Join-Path $PSScriptRoot "data"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

$output = [PSCustomObject]@{
    date = $today
    generated_at = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    total_news = $totalNews
    categories = $outputCategories
}

$jsonPath = Join-Path $dataDir "$today.json"
$jsonStr = $output | ConvertTo-Json -Depth 10
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($jsonPath, $jsonStr, $utf8NoBom)
Write-Host "[保存] $jsonPath"

# 更新索引
$indexPath = Join-Path $dataDir "index.json"
$dates = @()
if (Test-Path $indexPath) {
    # 用正则提取日期，避免 ConvertFrom-Json 在 PS7 下把单元素数组退化为字符串
    $rawIndex = Get-Content $indexPath -Raw -Encoding utf8
    foreach ($m in [regex]::Matches($rawIndex, '\d{4}-\d{2}-\d{2}')) {
        if ($dates -notcontains $m.Value) { $dates += $m.Value }
    }
}
if ($dates -notcontains $today) { $dates += $today }
$dates = @($dates | Sort-Object -Descending)

# 只保留最近 7 天：超期的数据文件直接删除，索引同步剔除
$cutoff = (Get-Date).Date.AddDays(-6)
$keptDates = @()
foreach ($d in $dates) {
    $parsed = $null
    try { $parsed = [datetime]::ParseExact($d, 'yyyy-MM-dd', $null) } catch { $parsed = $null }
    if ($parsed -and $parsed -ge $cutoff) {
        $keptDates += $d
    } else {
        $oldFile = Join-Path $dataDir "$d.json"
        if (Test-Path $oldFile) {
            Remove-Item $oldFile -Force
            Write-Host "[清理] 删除过期数据 $d.json"
        }
    }
}
$dates = @($keptDates | Sort-Object -Descending)

# 确保输出为JSON数组（每个元素带引号）
$quotedDates = $dates | ForEach-Object { '"' + $_ + '"' }
$indexJson = "[" + ($quotedDates -join ',') + "]"
[System.IO.File]::WriteAllText($indexPath, $indexJson, $utf8NoBom)
Write-Host "[索引] 更新 $indexPath，共 $($dates.Count) 天"

Write-Host "`n[处理完成] 去重分类后共 $totalNews 条新闻"
foreach ($cat in $outputCategories.Keys) {
    Write-Host "  - ${cat}: $($outputCategories[$cat].Count) 条"
}

Write-Host "`n============================================================"
Write-Host "采集完成！"
Write-Host "============================================================"
