# tranexecl
golang execl 处理,用来拆解订单生成新的 execl 表格方便导入.

### build
- ```go get github.com/tealeg/xlsx ```
- 路径为 window 路径,请按需修改
- win32``` CGO_ENABLED=0 GOOS=windows GOARCH=386 go build ```
- win64 ```CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build ```
