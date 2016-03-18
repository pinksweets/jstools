// �ύX�_��������Ȃ�Excel�̍����`�F�b�N�Ƀh�E�]
(function() {
	var diffMode = 1, // 1:�V�[�g����v�Ŕ�r�A2:�V�[�g��index��v�Ŕ�r
		xlCellTypeLastCell = 11, // �g��ꂽ�Z���͈͓��̍Ō�̃Z��
		excelApp = new ActiveXObject('Excel.Application'),
		shell = new ActiveXObject('WScript.Shell'),
		fs = new ActiveXObject('Scripting.FileSystemObject'),
		require=function(src) {
			shell.CurrentDirectory = WScript.ScriptFullName.replace(WScript.ScriptName,'');
			return eval(fs.OpenTextFile(src).ReadAll());
		},
		message = require('message.js'),
		saveFile,
		diffExcel = function() {
			var fileName1, filePath1, fileName2, filePath2, sheetCount1, sheetCount2, book1, book2, sheet1, sheet2, 
				maxRow, maxCol, value1, value2, sheetNum = 1, savePath, ret=true, 
				FileFilter = '�G�N�Z���t�@�C��(*.xls),*.xls';
			try{
				WScript.Echo('2�ԖڂɑI�񂾃t�@�C������ɕύX�̂������Z�����}�[�N���ĕʖ��ۑ����܂�');
				if(!(filePath1 = excelApp.Application.GetOpenFileName(FileFilter)) || !(filePath2 = excelApp.Application.GetOpenFileName(FileFilter))) {
					ret = false;
				} else {
					// ��r����Excel�P
					excelApp.Workbooks.Open(filePath1, false, true);
					book1 = excelApp.ActiveWorkbook;
					fileName1 = book1.name;
					sheetCount1 = book1.Worksheets.Count;
					// ��r����Excel�Q
					excelApp.Workbooks.Open(filePath2, false, true);
					book2 = excelApp.ActiveWorkbook;
					fileName2 = book2.name;
					sheetCount2 = book2.Worksheets.Count;
					for(var i=1;i<=sheetCount1;i++){
						sheet1 = book1.Worksheets(i);
						var name1=sheet1.Name,
							name2;
						for(var j=1;j<=sheetCount2;j++){
							sheet2 = book2.Worksheets(j);
							name2 = sheet2.Name;
							if(name1==name2) {
								break;
							}
						}
						if(name1!=name2) {
							continue;
						}
						maxRow = sheet1.Cells.SpecialCells(xlCellTypeLastCell).Row;
						maxCol = sheet1.Cells.SpecialCells(xlCellTypeLastCell).Column;
						message.rewrite(name1+'�̔�r��', 0);
						for(var Row = 1;Row <= maxRow; Row++) {
							for(var Col = 1;Col <= maxCol; Col++) {
								value1 = sheet1.Cells(Row, Col).Value;
								value2 = sheet2.Cells(Row, Col).Value;
								value1 || (value1='');
								value2 || (value2='');
								if(value1 != value2) {
									markSheet(sheet2.Cells(Row, Col), 7, value1);
								}
							}
						}
						sheetNum++;
					}
					if(ret){
						saveFile = WScript.ScriptFullName.replace(WScript.ScriptName,'')+'diff_'+fileName2;
						book2.SaveAs(saveFile);
					}
				}
			} catch(e) {
				WScript.Echo(e.message);
				ret = false;
			} finally {
				book1 && (book1.Close(false));
				book2 && (book2.Close(false));
				return ret;
			}
		},
		markSheet = function(cell, color, comVal) {
			cell.Interior.ColorIndex = color;
			if(!cell.Comment) {
				cell.addComment();
			}
			if(cell.Comment) {
				comVal || (comVal='');
				cell.Comment.text('���̒l:'+comVal);
			}
		};
	try {
		if(diffExcel()){ 
			message.rewrite('�����`�F�b�N���ʂ��t�@�C���F�u'+saveFile+'�v�ɕۑ����܂����B');
		} else {
			WScript.Echo('�����𒆒f���܂����B');
		}
	} catch(e) {
		WScript.Echo(e.message);
	} finally {
		excelApp && (excelApp.Quit());
	}
})();
