// IE�𗧂��グ�ăc�[��������ʒm���郁�b�Z�[�W��\�����܂��B
// ���̂悤�Ɏg���Ă��������B
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
		title = '�ʒm���b�Z�[�W';
	return {// �����������F�^�C�g���Ə������b�Z�[�W�̎w�肪�����ꍇ�͌Ăяo�����Ȃ��Ă��悢
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
			// ���b�Z�[�W�̒ǉ�
			add : function(message) {
				this.ie || (this.init());
				this.ie.document.body.innerHTML+=message+'<br/>';
			},
			// ���b�Z�[�W�̏�������
			rewrite : function(message) {
				this.ie || (this.init());
				this.ie.document.body.innerHTML=message+'<br/>';
			},
			// �I������
			end : function() {
				this.ie && (this.ie.Quit());
			}
		};
})();
