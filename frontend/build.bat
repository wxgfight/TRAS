@echo off
REM 清理旧的依赖
rmdir /s /q node_modules
rmdir /s /q dist
del /f /q package-lock.json

REM 安装依赖
echo 安装依赖中...
"C:\Program Files\nodejs\npm" install --force

REM 构建项目
echo 构建项目中...
"C:\Program Files\nodejs\node.exe" .\node_modules\vite\bin\vite.js build

REM 检查构建结果
echo 构建完成，检查dist目录：
dir dist
