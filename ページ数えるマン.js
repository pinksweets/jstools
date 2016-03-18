(function() {
	var FATT_NORMAL               =   0,  // 標準ファイル
		FATT_READONLY             =   1,  // 読み取り専用ファイル(取得と設定可)
		FATT_HIDDEN               =   2,  // 隠しファイル(取得と設定可)
		FATT_SYSTEM               =   4,  // システムファイル(取得と設定可)
		FATT_VOLUME               =   8,  // ディスクドライブボリュームラベル(取得のみ可)
		FATT_DIRECTORY            =  16,  // ディレクトリ(取得のみ可)
		FATT_ARCHIVE              =  32,  // アーカイブファイル(取得と設定可)
		FATT_ALIAS                =  64,  // ショートカットファイル(取得のみ可)
		FATT_COMPRESSED           = 128,  // 圧縮ファイル(取得のみ可)
		FORREADING                =   1,  // 読み取り専用
		FORWRITING                =   2,  // 書き込み専用
		FORAPPENDING              =   8,  // 追加書き込み
		TRISTATE_TRUE             =  -1,  // Unicode
		TRISTATE_FALSE            =   0,  // ASCII
		TRISTATE_USEDEFAULT       =  -2,  // システムデフォルト
		xlPageBreakPreview        =   2,
		wdNumberOfPagesInDocument =   4,
		excelApp = new ActiveXObject('Excel.Application'),
		wordApp = new ActiveXObject('Word.Application'),
		shell = new ActiveXObject('WScript.Shell'),
		fs = new ActiveXObject('Scripting.FileSystemObject'),
		outfile,
		getPageCountExcel = function(filepath, name) {
			var hpage, vpage, p = 0, sCnt, readSheet, book;
			try{
				excelApp.Workbooks.Open(filepath, false, true);
				//excelApp.Visible = true;
				book = excelApp.ActiveWorkbook;
				for(sCnt = 1;book.Worksheets.Count >= sCnt;sCnt++) {
					readSheet = book.Worksheets(sCnt);
					readSheet.Activate();
					//readSheet.Select();
					// Excel2007以降はPageSetup.Pages.Count推奨
					p += excelApp.ExecuteExcel4Macro('get.document(50)');
				}
				book.Close(false);
			} catch(e) {
				outfile.Write(filepath + ' / getPageCountExcel:' + e.message + '\r\n');
				book && book.Close(false);
			}
			return p;
		},
		getPageCountWord = function(filepath, name) {
			var doc, page = 0;
			try{
				wordApp.Documents.Open(filepath, false, true);
				doc = wordApp.Documents(name);
				page = wordApp.Selection.Information(wdNumberOfPagesInDocument);
				doc.Close(false);
			} catch(e) {
				outfile.Write(filepath + ' / getPageCountWord:' + e.message + '\r\n');
				doc && doc.Close(false);
			}
			return page;
		},
		listUp = function(fullPath, pages) {
			var fileList, folderList, subFolderList, folder, page, name, path;
			pages || (pages = 0);
			try{
				folder = fs.GetFolder(fullPath);
				if((folder.Attributes & FATT_NORMAL) || (folder.Attributes & FATT_DIRECTORY)) {
					fileList = folder.Files;
					var fc = new Enumerator(folder.Files);
					for (; !fc.atEnd(); fc.moveNext()) {
						name = fc.item().name;
						path = fc.item().path;
						if(name.match(/.+\.xls$/)) {
							page = getPageCountExcel(path, name);
							outfile.Write(path + '\t' + name + '\t' + page + '\r\n');
							pages += page;
						} else if(name.match(/.+\.doc$/)) {
							page = getPageCountWord(path, name)
							outfile.Write(path + '\t' + name + '\t' + page + '\r\n');
							pages += page;
						}
					}
					fc = new Enumerator(folder.SubFolders);
					for (; !fc.atEnd(); fc.moveNext()) {
						pages = listUp(fc.item().Path, pages);
					}
				}
			} catch(e) {
				outfile.Write('listUp:' + e.message);
			}
			return pages;
		};

	try{
		if(WScript.Arguments.length == 0) {
			WScript.Echo('ページ数を数えるドキュメントが入ったフォルダをドラッグ＆ドロップしてください');
		} else {
			shell.CurrentDirectory = WScript.ScriptFullName.replace(WScript.ScriptName,'');
			outfile = fs.OpenTextFile('./pages.log', FORWRITING, true, TRISTATE_FALSE)
			outfile.Write('フルパス\tファイル名\tページ数\r\n');
			var total = listUp(WScript.Arguments(0));
			outfile.Write('総ページ数：' + total);
			WScript.Echo('総ページ数：' + total);
		}
	} catch(e) {
		outfile.Write('もと:' + e.message);
	}
	excelApp && (excelApp.Quit());
	wordApp && (wordApp.Quit());
	outfile && (outfile.Close());
})();
