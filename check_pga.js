// ���A���^�C����PGA�g�p�ʂ��Ď����A5M�𒴂���ꍇ�͉�ʏo�͂��܂��B
// �I��������ꍇ�́A�ʒm���b�Z�[�W�E�B���h�E����邩�A
// ���c�[���Ɠ����t�H���_��stop.txt���쐬���Ă��������B
(function(){
	var execParams = {
			// �ڍ׏����e�L�X�g�o�͂���ꍇ�Atrue
			"oracleCheck" : false, 
			// DB�ڑ��ݒ�
			"connStr" : "Provider=OraOLEDB.Oracle; Data Source=ex; User Id=scotte; Password=tiger;",
			// �Ď��Ԋu�i�~���b�j
			"sleep_msec" : 5000
		},
		FORREADING	  = 1,	// �ǂݎ���p
		FORWRITING	  = 2,	// �������ݐ�p
		FORAPPENDING	= 8,	// �ǉ���������
		TRISTATE_TRUE	   = -1,   // Unicode
		TRISTATE_FALSE	  =  0,   // ASCII
		TRISTATE_USEDEFAULT = -2,   // �V�X�e���f�t�H���g
		shell = new ActiveXObject('WScript.Shell'),
		fs = new ActiveXObject('Scripting.FileSystemObject'),
		// �O���X�N���v�g�ǂݍ���
		require=function(src) {
			var ret;
			shell.CurrentDirectory = WScript.ScriptFullName.replace(WScript.ScriptName,'');
			try {
				ret = eval(fs.OpenTextFile(src, FORREADING).ReadAll());
			} catch(e) {
				WScript.Echo(e.name+' in require function.\n�@message:'+e.message);
			} finally{
				return ret;
			}
		},
		message = require('message.js'),
		conn = new ActiveXObject("ADODB.Connection"),
		rs = new ActiveXObject("ADODB.Recordset"),
		outputQuerys = {
				"�Z�b�V�������" : "select * from v$session",
				"PGA�̈�g�p�ʂ������Z�b�V����" : "select s.value, e.sid, e.username, e.program, e.process, e.last_call_et from v$sesstat s, v$session e, v$statname n where s.statistic# = n.statistic# and s.sid = e.sid and n.name = 'session pga memory' order by s.value desc, last_call_et asc",
				"�Z�b�V�����ڍ׏��" : "select s.sid,s.serial#,s.username,s.status,s.program,s.module,s.osuser,s.machine,s.command,s.sql_id,s.prev_sql_id, s.prev_exec_start,s.plsql_object_id, p.pga_used_mem,p.pga_alloc_mem,p.pga_freeable_mem,p.pga_max_mem,s.logon_time,sysdate,s.last_call_et, s.event,s.wait_class,s.wait_time,s.seconds_in_wait,s.state from v$session s,v$process p where s.paddr=p.addr order by p.pga_alloc_mem desc",
				"�r�p�k���擾(ACTIVE �܂��� �o�ߎ��Ԃ������r�p�k)" : "SELECT (sysdate - s.sql_exec_start) * 86400 as EXE_TIME, s.sid,s.USERNAME,q.* FROM( SELECT sql_id, sql_fulltext, address, hash_value, parse_calls, executions, buffer_gets, disk_reads, buffer_gets/executions buffer_per_run, disk_reads/executions disk_per_run, cpu_time, elapsed_time, elapsed_time/1000000/executions as AVG_TIME FROM v$sql WHERE executions > 0 and sql_text not like '%FROM v$sql%') q INNER JOIN v$session s ON s.sql_id = q.sql_id WHERE (s.status = 'ACTIVE' or (sysdate - s.sql_exec_start) * 86400 > 5) and s.program != 'OMS' and rownum < 41 ORDER BY 1 DESC",
				"���s�v��(ACTIVE �܂��� �o�ߎ��Ԃ������r�p�k)" : "select sid, p.sql_id, cardinality, bytes, cost, time , lpad(' ', depth) || operation || ' ' || options || ' ' || object_name as \"OPERATION\" from v$sql_plan p inner join v$sql q on q.sql_id = p.sql_id inner join v$session s on s.sql_id = p.sql_id where (s.status = 'ACTIVE' or (sysdate - s.sql_exec_start) * 86400 > 5) and s.program != 'OMS' and q.sql_text not like '%FROM v$sql%' order by sid, timestamp, id"
			},
		pgaOverSql = [
			"SELECT s.sid, to_char(sysdate,'YYYY/MM/DD hh24:mi:ss') as ymd, to_char(s.value, 'fm999,999,999,999') as pga, e.username as un, q.sql_fulltext ",
			"FROM v$sesstat s, v$session e, v$statname n, v$sql q ",
			"WHERE s.statistic# = n.statistic# AND s.sid = e.sid AND ",
			"  n.name = 'session pga memory' AND e.sql_id = q.sql_id AND ",
			"  s.value >= 5000000 AND e.username = '" + execParams["owner"] + "' ",
			"ORDER BY s.value DESC, last_call_et ASC"],
		sysdateStr = function() {
			var _d = new Date();
			return [_d.getFullYear(), ('0'+(_d.getMonth() + 1)).slice(-2), ('0'+_d.getDate()).slice(-2), ('0'+_d.getHours()).slice(-2),('0'+_d.getMinutes()).slice(-2),('0'+_d.getSeconds()).slice(-2)].join("");
		},
		reportOracleCheck = function() {
			var output = fs.OpenTextFile(sysdateStr()+'.csv', FORWRITING, true, TRISTATE_FALSE), 
				v = new Array(), i, fsize, title = new Array();
			for (var key in outputQuerys) {
				output.Write(key + '\r\n');
				rs.Open(outputQuerys[key], conn);
				fsize = rs.Fields.Count;
				title.length = 0;
				for (i = 0;fsize > i; i++) {
					title.push(rs.Fields(i).Name);
				}
				output.Write(title.join(',') + '\r\n');
				while (!rs.EOF) {
					v.length = 0;
					for (i = 0;fsize > i; i++) {
						v.push(""+rs.Fields(i));
					}
					output.Write(v.join(',') + '\r\n');
					rs.MoveNext();
				}
				rs.Close();
				message.add(key + " fields=" + fsize + " / " + v.length);
			}
			output.Close();
		},
		peakPga = {};
	try {
		message.add('�������Ď��J�n');
		conn.Open(execParams["connStr"]);
		while (!fs.FileExists('stop.txt')) {
			execParams["oracleCheck"] && (reportOracleCheck());
			rs.Open(pgaOverSql.join(' '), conn);
			while (!rs.EOF) {
				message.add(rs.Fields('ymd') + ' / ' + rs.Fields('un') + ' / sid=' + rs.Fields('sid') + ' / pga=' + rs.Fields('pga'));
				0+peakPga[rs.Fields('sql_fulltext')] < rs.Fields('pga') && (peakPga[rs.Fields('sql_fulltext')] = rs.Fields('pga'));
				rs.MoveNext();
			}
			rs.Close();
			WScript.Sleep(execParams["sleep_msec"]);
		}
	} catch(e) {
		//WScript.Echo(e.name+' in main logic.\n�@message:'+e.message);
	} finally {
		var output = fs.OpenTextFile('checked_sql.txt', FORWRITING, true, TRISTATE_FALSE);
		for (sql in peakPga) {
			output.Write("pga="+peakPga[sql]+" / sql="+sql);
		}
		output.Close();
		conn.Close();
		WScript.Echo("check_pga.js���I�����܂���");
	}
})();