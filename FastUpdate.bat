@echo off
set /p url="Your repo name:"
git init
git remote add origin git@github.com:buger404/%url%.git
git remote set-url --add origin git@gitee.com:buger404/%url%.git
git add .
git commit -m "first commit"
git push -u origin master