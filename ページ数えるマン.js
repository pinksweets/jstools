(function() {
	var FATT_NORMAL               =   0,  // �W���t�@�C��
		FATT_READONLY             =   1,  // �ǂݎ���p�t�@�C��(�擾�Ɛݒ��)
		FATT_HIDDEN               =   2,  // �B���t�@�C��(�擾�Ɛݒ��)
		FATT_SYSTEM               =   4,  // �V�X�e���t�@�C��(�擾�Ɛݒ��)
		FATT_VOLUME               =   8,  // �f�B�X�N�h���C�u�{�����[�����x��(�擾�̂݉�)
		FATT_DIRECTORY            =  16,  // �f�B���N�g��(�擾�̂݉�)
		FATT_ARCHIVE              =  32,  // �A�[�J�C�u�t�@�C��(�擾�Ɛݒ��)
		FATT_ALIAS                =  64,  // �V���[�g�J�b�g�t�@�C��(�擾�̂݉�)
		FATT_COMPRESSED           = 128,  // ���k�t�@�C��(�擾�̂݉�)
		FORREADING                =   1,  // �ǂݎ���p
		FORWRITING                =   2,  // �������ݐ�p
		FORAPPENDING              =   8,  // �ǉ���������
		TRISTATE_TRUE             =  -1,  // Unicode
		TRISTATE_FALSE            =   0,  // ASCII
		TRISTATE_USEDEFAULT       =  -2,  // �V�X�e���f�t�H���g
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
					// Excel2007�ȍ~��PageSetup.Pages.Count����
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
			WScript.Echo('�y�[�W���𐔂���h�L�������g���������t�H���_���h���b�O���h���b�v���Ă�������');
		} else {
			shell.CurrentDirectory = WScript.ScriptFullName.replace(WScript.ScriptName,'');
			outfile = fs.OpenTextFile('./pages.log', FORWRITING, true, TRISTATE_FALSE)
			outfile.Write('�t���p�X\t�t�@�C����\t�y�[�W��\r\n');
			var total = listUp(WScript.Arguments(0));
			outfile.Write('���y�[�W���F' + total);
			WScript.Echo('���y�[�W���F' + total);
		}
	} catch(e) {
		outfile.Write('����:' + e.message);
	}
	excelApp && (excelApp.Quit());
	wordApp && (wordApp.Quit());
	outfile && (outfile.Close());
})();
