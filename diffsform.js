//----------------------------------------------------------------------------//
//------------------------- Сверка расхождений f915. -------------------------//
//----------------------------------------------------------------------------//
function SaveToFile(fname, arr)
{
	var objFS, objFile, i;

	objFS = new ActiveXObject('Scripting.FileSystemObject');
	objFile = objFS.CreateTextFile(fname);
	for(i in arr)
	{
		objFile.WriteLine(arr[i]);
	}
	objFile.Close();

	return 0;
}

function SQL2ARR(obj)
{
	var objRecords, objFields, Records, Fields;

	objRecords = obj;

	Records = new Array();
	if((!(objRecords.EOF)))
	{
		for((objRecords.MoveFirst()); (!(objRecords.EOF)); (objRecords.MoveNext()))
		{
			if((Records.length == 0))
			{
				objFields = new Enumerator(objRecords.Fields);
				Fields = new Array();
				if((!(objFields.atEnd())))
				{
					for((objFields.moveFirst()); (!(objFields.atEnd())); (objFields.moveNext()))
					{
						Fields.push(objFields.item().Name);
					}
				}
				Records.push(Fields);
			}
			if((Records.length != 0))
			{
				objFields = new Enumerator(objRecords.Fields);
				Fields = new Array();
				if((!(objFields.atEnd())))
				{
					for((objFields.moveFirst()); (!(objFields.atEnd())); (objFields.moveNext()))
					{
						Fields.push(objFields.item().Value);
					}
				}
				Records.push(Fields);
			}
		}
	}

	return Records;
}

function ExecSQL(str, sql)
{
	var Result, objConn;

	objConn = new ActiveXObject('ADODB.Connection');
	objConn.Open(str);
	Result = SQL2ARR(objConn.Execute(sql));
	objConn.Close();

	return Result;
}

function getKeyVal(arr)
{
	var Result, str, tmp, key, val, i, k;

	Result = new Array();

	for(i in arr)
	{
		key = new Array();
		val = new Array();
		for(k in arr[i])
		{
			str = arr[0][k].split(':');
			if((i == 0))
			{
				tmp = str[1];
			}
			if((i != 0))
			{
				tmp = arr[i][k];
			}
			switch(str[0])
			{
				case 'KEY':
				{
					key.push(tmp);
				}
				break;
				case 'VALUE':
				{
					val.push(tmp);
				}
				break;
			}
		}
		tmp = new Array();
		tmp.push(key.join(';'));
		tmp.push(val.join(';'));
		Result.push(tmp.join(':'));
	}

	return Result;
}

function getHashSize(arr)
{
	var Result, i;

	Result = new Number();

	for(i in arr)
	{
		if((arr[i]))
		{
			Result += 1;
		}
	}

	return Result;
}

function getCustomSTR(type)
{
	var Result;

	switch(type)
	{
		case 'FRONT':
		{
			Result = new String('Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP) (HOST = chalna.cgs.sbrf.ru) (PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = uzb_pcod)));User ID=wsback;Password=wsback');
		}
		break;
		case 'BACK':
		{
			Result = new String('Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP) (HOST = osma.cgs.sbrf.ru) (PORT = 1521)) (CONNECT_DATA = (SID = dbdpc)));User ID=wsback;Password=wsback');
		}
		break;
	}

	return Result;
}

function getCustomSQL(type, vsp)
{
	var Result, arr;

	switch(type)
	{
//		BASE
		case 'FORM':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.branchno || \'/\' || t.office as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.outotal), 2) as "VALUE:Сумма", trunc(sum(t.oures), 0) as "VALUE:Количество" from operday.form915 t where(((t.day, t.id_mega) in ((select max(t.day), t.id_mega from operday.form915 t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.id_mega) in ((52)))) group by(t.id_mega))))) having((not((trunc(sum(t.outotal), 2), trunc(sum(t.oures), 0)) in ((0.00, 0))))) group by(t.day, t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
		case 'TOTAL':
		{
			Result = new String('select to_char(trunc(sysdate, \'YYYY\'), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.branchno || \'/\' || t.office as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.cash), 2) as "VALUE:Сумма", trunc(sum(t.cnt), 0) as "VALUE:Количество" from depo_stat.officetotal t where((not((t.kind) in ((10)))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.cash), 2), trunc(sum(t.cnt), 0)) in ((0.00, 0))))) group by(sysdate, t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
//		TURN
		case 'ESK':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.payoffcash), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.eskturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'MCG':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.payoffcash), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.mcgturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'SYNC':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.payoffcash), 2) as "VALUE:Расход", trunc(sum(t.opencnt), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.syncturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(t.opencnt), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'OFFICE':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.branchno || \'/\' || t.office as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.branchno || \'/\' || t.office as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.incashin + t.offcashin + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.incashou + t.offcashou), 2) as "VALUE:Расход", trunc(sum(t.opencnt), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.officeturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.branchno, t.office, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incashin + t.offcashin + t.prcntcash), 2), trunc(sum(t.incashou + t.offcashou), 2), trunc(sum(t.opencnt), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.branchno, t.office, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'OFFCASH':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.payoffcash), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.offcashturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'MOFFICE':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.pairoffcashin + t.pairprcntcash), 2) as "VALUE:Приход", trunc(sum(t.pairoffcashou), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.mofficeturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.pairoffcashin + t.pairprcntcash), 2), trunc(sum(t.pairoffcashou), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'SOFFICE':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.soffice as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.cash), 2) as "VALUE:Приход", trunc(sum(t.cash), 2) as "VALUE:Расход", trunc(sum(t.cnt), 0) as "VALUE:Открыто", trunc(sum(t.closecnt), 0) as "VALUE:Закрыто" from depo_stat.officesplitturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.cash), 2), trunc(sum(t.cash), 2), trunc(sum(t.cnt), 0), trunc(sum(t.closecnt), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.soffice, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'CAPITAL':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.incash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.outcash), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(0), 0) as "VALUE:Закрыто" from depo_stat.capitalturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incash + t.prcntcash), 2), trunc(sum(t.outcash), 2), trunc(sum(0), 0), trunc(sum(0), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'PROLONG':
		{
			Result = new String('select to_char(trunc(t.day), \'DD.MM.YYYY\') as "KEY:Дата", t.branchno || \'/\' || t.office as "KEY:ВСП опер.", t.dbranchno || \'/\' || t.doffice as "KEY:ВСП счет.", t.kind || \'.\' || t.subkind as "KEY:Вклад", t.currency as "KEY:Валюта", t.account as "KEY:Баланс", t.dbranchno || \'/\' || t.doffice as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(t.incash + t.prcntcash), 2) as "VALUE:Приход", trunc(sum(t.outcash), 2) as "VALUE:Расход", trunc(sum(0), 0) as "VALUE:Открыто", trunc(sum(0), 0) as "VALUE:Закрыто" from depo_stat.prolongturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incash + t.prcntcash), 2), trunc(sum(t.outcash), 2), trunc(sum(0), 0), trunc(sum(0), 0)) in ((0, 0, 0, 0))))) group by(t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		} 
		break;
//		ACCOUNT
		case 'ACCOUNT':
		{
			Result = new String('select substr(t.printableno, 1, 8) || \'x\' || substr(t.printableno, 10, 11) as "KEY:Номер счета", t.account as "KEY:Баланс", to_char(trunc(t.opday), \'DD.MM.YYYY\') as "VALUE:Дата опер.", t.branchno || \'/\' || t.office as "VALUE:ВСП счет.", t.kind || \'.\' || t.subkind as "VALUE:Вклад", t.currency as "VALUE:Валюта", trunc(sum(decode(t.opno, 0, 1, t.opno)), 0) as "VALUE:Номер опер.", trunc(sum(t.opcash), 2) as "VALUE:Сумма", trunc(sum(t.balance), 2) as "VALUE:Остаток", trunc(sum(decode(t.state, 4, 0, 5, 0, t.state)), 0) as "VALUE:Статус" from deposit.deposit t where(((t.opday) between (to_date(\'01.01.1600\', \'DD.MM.YYYY\')) and (to_date(sysdate, \'DD.MM.YYYY\'))) and ((t.branchno, t.office, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) group by(t.printableno, t.account, t.opday, t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
	}
	vsp = vsp.replace('/', ';');
	vsp = vsp.replace('.', ';');
	arr = vsp.split(';');
	Result = Result.replace('[BRANCH]', arr[1-1]);
	Result = Result.replace('[OFFICE]', arr[2-1]);
	Result = Result.replace('[KIND]', arr[3-1]);
	Result = Result.replace('[SUBKIND]', arr[4-1]);
	Result = Result.replace('[CURRENCY]', arr[5-1]);

	return Result;
}

function getCustomDiffs(type, vsp)
{
	var Result, DB, sql, str, tmp, arr, key, val, hd, kv, i, k;

	DB = new Array('FRONT', 'BACK');

	arr = new Array();
	for(i in DB)
	{
		str = getCustomSTR(DB[i]);
		sql = getCustomSQL(type, vsp);
		tmp = getKeyVal(ExecSQL(str, sql));
		for(k in tmp)
		{
			if((k == 0))
			{
				hd = tmp[k].split(':');
			}
			if((k != 0))
			{
				kv = tmp[k].split(':');
				if((!(arr[kv[1-1]])))
				{
					arr[kv[1-1]] = new Array();
				}
				arr[kv[1-1]][i] = kv[2-1];
			}
		}
	}

	Result = new Array();
	for(k in arr)
	{
		tmp = new Array();
		for(i in DB)
		{
			if((!(arr[k][i])))
			{
				arr[k][i] = new String();
			}
			tmp[arr[k][i]] = k + i;
		}
		if((getHashSize(tmp) != 1))
		{
			key = new Array();
			tmp = new Array();
			tmp.push('Тип');
			tmp.push(hd[1-1]);
			key.push(tmp.join(';'));
			tmp = new Array();
			tmp.push(type);
			tmp.push(k);
			key.push(tmp.join(';'));
			val = new Array();
			tmp = new Array();
			tmp.push('Система');
			tmp.push(hd[2-1]);
			val.push(tmp.join(';'));
			for(i in DB)
			{
				tmp = new Array();
				tmp.push(DB[i]);
				tmp.push(arr[k][i]);
				val.push(tmp.join(';'));
			}
			Result[key.join(':')] = val.join(':');
		}
	}

	return Result;
}

function getCommonDiffs(type, vsp)
{
	var Result, types, diffs, i, k;

	switch(type)
	{
		case 'BASE':
		{
			types = new Array('FORM', 'TOTAL');
		}
		break;
		case 'TURN':
		{
			types = new Array('ESK', 'MCG', 'SYNC', 'OFFICE', 'OFFCASH', 'MOFFICE', 'SOFFICE', 'CAPITAL', 'PROLONG');
		}
		break;
		case 'ACCOUNT':
		{
			types = new Array('ACCOUNT');
		}
		break;
	}

	Result = new Array();
	for(k in types)
	{
		diffs = getCustomDiffs(types[k], vsp);
		for(i in diffs)
		{
			Result[i] = diffs[i];
		}
	}

	return Result;
}

function getVSPList(diffs)
{
	var Result, arr, tmp, str, i, k;

	Result = new Array();
	for(i in diffs)
	{
		arr = i.split(':');
		arr = arr[2-1].split(';');
		arr.reverse();
		arr.pop();
		arr.pop();
		arr.reverse();
		Result.push(arr.join(';'));
	}

	return Result;
}

function makeHTML(diffs, FileName)
{
	var content, values, items, i, k, m;

	content = new Array();
	content.push('<html>');
	content.push('<head>');
	content.push('<title>');
	content.push('Сверка расхождений f915.');
	content.push('</title>');
	content.push('</head>');
	content.push('<body>');
	for(i in diffs)
	{
		content.push('<table>');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=0 cellspacing=0>');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=3 cellspacing=0>');
		values = i.split(':');
		for(k in values)
		{
			content.push('<tr>');
			items = values[k].split(';');
			for(m in items)
			{
				content.push('<th>');
				content.push(items[m]);
				content.push('</th>');
			}
			content.push('</tr>');
		}
		content.push('</table>');
		content.push('</td>');
		content.push('</tr>');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=3 cellspacing=0>');
		values = diffs[i].split(':');
		for(k in values)
		{
			content.push('<tr>');
			items = values[k].split(';');
			for(m in items)
			{
				content.push('<th>');
				content.push(items[m]);
				content.push('</th>');
			}
			content.push('</tr>');
		}
		content.push('</table>');
		content.push('</td>');
		content.push('</tr>');
		content.push('</table>');
		content.push('</td>');
		content.push('</tr>');
		content.push('</table>');
	}
	content.push('</body>');
	content.push('</html>');

	SaveToFile(FileName, content);

	return 0;
}

function Main()
{
	var diffs, tmp, vsp, i, k;

	vsp = new Array('');
	diffs = new Array();
	for(i in vsp)
	{
		tmp = getCommonDiffs('BASE', vsp[i]);
		for(k in tmp)
		{
			diffs[k] = tmp[k];
		}
	}
	makeHTML(diffs, './files/diffsfrm.html');

	vsp = getVSPList(diffs);
	diffs = new Array();
	for(i in vsp)
	{
		tmp = getCommonDiffs('TURN', vsp[i]);
		for(k in tmp)
		{
			diffs[k] = tmp[k];
		}
	}
	makeHTML(diffs, './files/diffstrn.html');

	vsp = vsp;
	diffs = new Array();
	for(i in vsp)
	{
		tmp = getCommonDiffs('ACCOUNT', vsp[i]);
		for(k in tmp)
		{
			diffs[k] = tmp[k];
		}
	}
	makeHTML(diffs, './files/diffsacc.html');

	return 0;
}


{
	Main();
}
