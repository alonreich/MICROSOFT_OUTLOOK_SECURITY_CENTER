Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & WScript.Arguments(0) & chr(34) & " --service --disable-gpu --disable-software-rasterizer --disable-gpu-compositing --disable-gpu-sandbox --disable-accelerated-2d-canvas --use-gl=disabled --disable-vulkan --disable-gpu-shader-disk-cache --disable-gpu-rasterization --no-sandbox", 0
Set WshShell = Nothing
