// �w�肵���f�B���N�g�������񂵂܂��B
// �����ɂ��A�t�@�C�����̃p�^�[���}�b�`���O�ɂ��i���݂ƃR�[���o�b�N�֐��̌Ăяo�����s���܂��B
// var fs = new ActiveXObject('Scripting.FileSystemObject'),
// require=function(src) {
// 	return eval(fs.OpenTextFile(src).ReadAll());
// },
// req = require('./fsworker.js'),
(function() {
	var FATT_NORMAL     =   0,  // �W���t�@�C��
		FATT_READONLY   =   1,  // �ǂݎ���p�t�@�C��(�擾�Ɛݒ��)
		FATT_HIDDEN     =   2,  // �B���t�@�C��(�擾�Ɛݒ��)
		FATT_SYSTEM     =   4,  // �V�X�e���t�@�C��(�擾�Ɛݒ��)
		FATT_VOLUME     =   8,  // �f�B�X�N�h���C�u�{�����[�����x��(�擾�̂݉�)
		FATT_DIRECTORY  =  16,  // �f�B���N�g��(�擾�̂݉�)
		FATT_ARCHIVE    =  32,  // �A�[�J�C�u�t�@�C��(�擾�Ɛݒ��)
		FATT_ALIAS      =  64,  // �V���[�g�J�b�g�t�@�C��(�擾�̂݉�)
		FATT_COMPRESSED = 128,  // ���k�t�@�C��(�擾�̂݉�)
		fs = new ActiveXObject('Scripting.FileSystemObject'),
		fullPath, pattern, fn;
	return {// �T������X�^�[�g�n�_�̃t�H���_���w�肵�܂�
			setFullPath : function(arg){fullPath = arg;return this;}, 
			// �T������t�@�C���̃p�^�[���}�b�`���w�肵�܂�
			setPattern : function(arg){pattern = arg;return this;}, 
			// �T������t�@�C���ɑ΂��鏈�����w�肵�܂�
			setFn : function(arg){fn = arg;return this;},
			// �ċA���čŉ��w�܂ŏ��񂵂ăp�^�[���Ɉ�v����t�@�C���ɑ΂��Ďw�肳�ꂽ���������s���܂��B
			work : function(next) {
				var folder, name, path, fc;
				if(!fullPath) {
					WScript.Echo('setFullPath�ŒT������X�^�[�g�n�_�̃t�H���_���w�肵�Ă�������');
					return false;
				}
				if(!pattern) {
					WScript.Echo('setPattern�ŒT������t�@�C���̃p�^�[���}�b�`���w�肵�Ă�������');
					return false;
				}
				if(!fn) {
					WScript.Echo('setFn�ŒT�����ʂɑ΂��鏈����ݒ肵�Ă�������');
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
