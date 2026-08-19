set x to POSIX file "/Users/apple/SocratesAI-codex/行业智能化第一性原理.pptx"
set y to "/Users/apple/SocratesAI-codex/行业智能化第一性原理.ppt"
tell application "/Applications/Microsoft PowerPoint.app"
	open x
	save active presentation in y
	close active presentation saving no
end tell
