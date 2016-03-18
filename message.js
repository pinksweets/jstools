// IEを立ち上げてツール内から通知するメッセージを表示します。
// ↓のように使ってください。
// var fs = new ActiveXObject('Scripting.FileSystemObject'),
// require=function(src) {
// 	return eval(fs.OpenTextFile(src).ReadAll());
// },
// req = require('./message.js'),
var tst = (function() {
	var waitIE = function() {
			while((this.ie.Busy) || (this.ie.readystate != 4)) {
				WScript.Sleep(100);
			}
		},
		title = '通知メッセージ';
	return {// 初期化処理：タイトルと初期メッセージの指定が無い場合は呼び出ししなくてもよい
			init : function(tit, msg) {
				this.ie = new ActiveXObject('InternetExplorer.Application');
				this.ie.Width = 400;
				this.ie.Height = 140;
				this.ie.AddressBar = false;
				this.ie.MenuBar = false;
				this.ie.ToolBar = false;
				this.ie.StatusBar = false;
				this.ie.Resizable = true;
				this.ie.Visible = true;
				this.ie.Navigate('about:blank');
				msg || (msg='');
				tit || (tit=title);
				waitIE();
				this.ie.document.tit=tit;
				if(msg) {
					this.ie.document.body.innerHTML=msg+'<br/>';
				}
			},
			// メッセージの追加
			add : function(message) {
				this.ie || (this.init());
				this.ie.document.body.innerHTML+=message+'<br/>';
			},
			// メッセージの書き換え
			rewrite : function(message) {
				this.ie || (this.init());
				this.ie.document.body.innerHTML=message+'<br/>';
			},
			// 終了処理
			end : function() {
				this.ie && (this.ie.Quit());
			}
		};
})();
