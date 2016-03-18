// 指定したディレクトリを巡回します。
// 引数により、ファイル名のパターンマッチングによる絞込みとコールバック関数の呼び出しを行います。
// var fs = new ActiveXObject('Scripting.FileSystemObject'),
// require=function(src) {
// 	return eval(fs.OpenTextFile(src).ReadAll());
// },
// req = require('./fsworker.js'),
(function() {
	var FATT_NORMAL     =   0,  // 標準ファイル
		FATT_READONLY   =   1,  // 読み取り専用ファイル(取得と設定可)
		FATT_HIDDEN     =   2,  // 隠しファイル(取得と設定可)
		FATT_SYSTEM     =   4,  // システムファイル(取得と設定可)
		FATT_VOLUME     =   8,  // ディスクドライブボリュームラベル(取得のみ可)
		FATT_DIRECTORY  =  16,  // ディレクトリ(取得のみ可)
		FATT_ARCHIVE    =  32,  // アーカイブファイル(取得と設定可)
		FATT_ALIAS      =  64,  // ショートカットファイル(取得のみ可)
		FATT_COMPRESSED = 128,  // 圧縮ファイル(取得のみ可)
		fs = new ActiveXObject('Scripting.FileSystemObject'),
		fullPath, pattern, fn;
	return {// 探査するスタート地点のフォルダを指定します
			setFullPath : function(arg){fullPath = arg;return this;}, 
			// 探査するファイルのパターンマッチを指定します
			setPattern : function(arg){pattern = arg;return this;}, 
			// 探査するファイルに対する処理を指定します
			setFn : function(arg){fn = arg;return this;},
			// 再帰して最下層まで巡回してパターンに一致するファイルに対して指定された処理を実行します。
			work : function(next) {
				var folder, name, path, fc;
				if(!fullPath) {
					WScript.Echo('setFullPathで探査するスタート地点のフォルダを指定してください');
					return false;
				}
				if(!pattern) {
					WScript.Echo('setPatternで探査するファイルのパターンマッチを指定してください');
					return false;
				}
				if(!fn) {
					WScript.Echo('setFnで探査結果に対する処理を設定してください');
					return false;
				}
				try{
					next || (next = fullPath);
					folder = fs.GetFolder(next);
					if(folder.Attributes & FATT_NORMAL || folder.Attributes & FATT_DIRECTORY) {
						fc = new Enumerator(folder.Files);
						while (!fc.atEnd()) {
							name = fc.item().name;
							path = fc.item().path;
							if(pattern.test(name)) {
								fn(path, name);
							}
							fc.moveNext();
						}
						fc = new Enumerator(folder.SubFolders);
						while (!fc.atEnd()) {
							this.work(fc.item().Path);
							fc.moveNext();
						}
					}
					return true;
				} catch(e) {
					WScript.Echo('fsworker.work:' + e.message + ' next:'+next);
					return false;
				}
			}
		};
})();
