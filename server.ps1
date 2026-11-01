# 简易本地HTTP服务器，用于测试前端页面
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:8080/")
$listener.Start()
Write-Host "Server running at http://localhost:8080/"
Write-Host "Press Ctrl+C to stop"

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $path = $request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }

        $filePath = Join-Path $PSScriptRoot $path.TrimStart("/")

        if (Test-Path $filePath) {
            $ext = [System.IO.Path]::GetExtension($filePath)
            $contentType = switch ($ext) {
                ".html" { "text/html; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".js"   { "application/javascript; charset=utf-8" }
                ".json" { "application/json; charset=utf-8" }
                default { "application/octet-stream" }
            }
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = $contentType
            $response.ContentLength64 = $bytes.Length
            $response.AddHeader("Cache-Control", "no-store, no-cache, must-revalidate")
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $response.StatusCode = 404
        }
        $response.Close()
    } catch {
        Write-Host "Error: $_"
    }
}
