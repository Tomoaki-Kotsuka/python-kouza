-- Python 講座の比較用 VBA(split_by_area)から呼ばれる補助スクリプト
-- Mac 版 Excel の VBA は日本語のフォルダ名を MkDir で作れないため、ここで作る
on MakeFolder(folderPath)
	do shell script "mkdir -p " & quoted form of folderPath
	return "ok"
end MakeFolder

on FolderExists(folderPath)
	try
		do shell script "test -d " & quoted form of folderPath
		return "1"
	on error
		return "0"
	end try
end FolderExists
