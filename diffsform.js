//----------------------------------------------------------------------------//
//-------------------------   f915. -------------------------//
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

function getCustomSTR(DB)
{
	var Result;

	switch(DB)
	{
		case 'FRONT':
		{
			Result = new String('Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP) (HOST = ) (PORT = )) (CONNECT_DATA = (SERVICE_NAME = )));User ID=;Password=');
		}
		break;
		case 'BACK':
		{
			Result = new String('Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP) (HOST = ) (PORT = )) (CONNECT_DATA = (SID = )));User ID=;Password=');
		}
		break;
		default:
		{
			throw new Error('Unsupported argument passed.');
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
			Result = new String('select to_char(trunc(min(t.day), \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", to_char(trunc(min(t.day), \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.branchno as "VAL: .", t.office as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.outotal), 2) as "VAL:", trunc(sum(t.oures), 0) as "VAL:" from operday.form915 t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.branchno, t.office, t.kind, t.subkind, t.currency, t.oures, t.outotal) in ((select t.branchno, t.office, t.kind, t.subkind, t.currency, t.oures, t.outotal from operday.form915 t where(((t.day, t.id_mega) in ((select max(t.day), t.id_mega from operday.form915 t where(((t.id_mega) in ((52)))) group by(t.id_mega))))))))) having((not((trunc(sum(t.outotal), 2), trunc(sum(t.oures), 0)) in ((0.00, 0))))) group by(t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
		case 'TOTAL':
		{
			Result = new String('select to_char(trunc(sysdate, \'YYYY\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", to_char(trunc(sysdate, \'YYYY\'), \'DD.MM.YYYY\') as "VAL:", t.branchno as "VAL: .", t.office as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.cash), 2) as "VAL:", trunc(sum(t.cnt), 0) as "VAL:" from depo_stat.officetotal t where((not((t.kind) in ((10)))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.cash), 2), trunc(sum(t.cnt), 0)) in ((0.00, 0))))) group by(sysdate, t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
//		TURN
		case 'ESK':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VAL:", trunc(sum(t.payoffcash), 2) as "VAL:", trunc(sum(0), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.eskturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'MCG':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VAL:", trunc(sum(t.payoffcash), 2) as "VAL:", trunc(sum(0), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.mcgturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'SYNC':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VAL:", trunc(sum(t.payoffcash), 2) as "VAL:", trunc(sum(t.opencnt), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.syncturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(t.opencnt), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'OFFICE':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.branchno as "KEY: .", t.office as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.assignday, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.branchno "VAL: .", t.office as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.incashin + t.offcashin + t.prcntcash), 2) as "VAL:", trunc(sum(t.incashou + t.offcashou), 2) as "VAL:", trunc(sum(t.opencnt), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.officeturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.branchno, t.office, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incashin + t.offcashin + t.prcntcash), 2), trunc(sum(t.incashou + t.offcashou), 2), trunc(sum(t.opencnt), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.assignday, t.branchno, t.office, t.branchno, t.office, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'OFFCASH':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.enrolcash + t.prcntcash), 2) as "VAL:", trunc(sum(t.payoffcash), 2) as "VAL:", trunc(sum(t.opencnt), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.offcashturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.enrolcash + t.prcntcash), 2), trunc(sum(t.payoffcash), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'MOFFICE':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.pairoffcashin + t.pairprcntcash), 2) as "VAL:", trunc(sum(t.pairoffcashou), 2) as "VAL:", trunc(sum(t.opcnt), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.mofficeturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.pairoffcashin + t.pairprcntcash), 2), trunc(sum(t.pairoffcashou), 2), trunc(sum(0), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'SOFFICE':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.soffice as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.cash), 2) as "VAL:", trunc(sum(t.cash), 2) as "VAL:", trunc(sum(t.cnt), 0) as "VAL:", trunc(sum(t.closecnt), 0) as "VAL:" from depo_stat.officesplitturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.cash), 2), trunc(sum(t.cash), 2), trunc(sum(t.cnt), 0), trunc(sum(t.closecnt), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.soffice, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'CAPITAL':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.incash + t.prcntcash), 2) as "VAL:", trunc(sum(t.outcash), 2) as "VAL:", trunc(sum(0), 0) as "VAL:", trunc(sum(0), 0) as "VAL:" from depo_stat.capitalturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incash + t.prcntcash), 2), trunc(sum(t.outcash), 2), trunc(sum(0), 0), trunc(sum(0), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		}
		break;
		case 'PROLONG':
		{
			Result = new String('select to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "KEY:", t.branchno as "KEY: .", t.office as "KEY: .", t.dbranchno as "KEY: .", t.doffice as "KEY: .", t.kind as "KEY:", t.subkind as "KEY:", t.currency as "KEY:", t.account as "KEY:", to_char(trunc(t.day, \'DD\'), \'DD.MM.YYYY\') as "VAL:", t.dbranchno "VAL: .", t.doffice as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(t.incash + t.prcntcash), 2) as "VAL:", trunc(sum(t.outcash), 2) as "VAL:", trunc(sum(0), 0) as "VAL:", trunc(sum(0), 0) as "VAL:" from depo_stat.prolongturn t where(((t.day) between (trunc(sysdate, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.dbranchno, t.doffice, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) having((not((trunc(sum(t.incash + t.prcntcash), 2), trunc(sum(t.outcash), 2), trunc(sum(0), 0), trunc(sum(0), 0)) in ((0.00, 0.00, 0, 0))))) group by(t.day, t.day, t.branchno, t.office, t.dbranchno, t.doffice, t.kind, t.subkind, t.currency, t.account)');
		} 
		break;
//		ACCOUNT
		case 'ACCOUNT':
		{
			Result = new String('select substr(t.printableno, 1, 8) || \'x\' || substr(t.printableno, 10, 11) as "KEY: ", t.account as "KEY:", to_char(trunc(t.opday, \'DD\'), \'DD.MM.YYYY\') as "VAL: .", t.branchno as "VAL: .", t.office as "VAL: .", t.kind as "VAL:", t.subkind as "VAL:", t.currency as "VAL:", trunc(sum(decode(t.opno, 0, 1, t.opno)), 0) as "VAL: .", trunc(sum(t.opcash), 2) as "VAL:", trunc(sum(t.balance), 2) as "VAL:", trunc(sum(decode(t.state, 4, 0, 5, 0, t.state)), 0) as "VAL:" from deposit.deposit t where(((t.opday) between (trunc(sysdate - 110000, \'YYYY\')) and (trunc(sysdate, \'DD\'))) and ((t.branchno, t.office, t.kind, t.subkind, t.currency) in (([BRANCH], [OFFICE], [KIND], [SUBKIND], [CURRENCY]))) and ((t.id_mega) in ((52)))) group by(t.printableno, t.account, t.opday, t.branchno, t.office, t.kind, t.subkind, t.currency)');
		}
		break;
		default:
		{
			throw new Error('Unsupported argument passed.');
		}
		break;
	}
	arr = vsp.split(';');
	Result = Result.replace('[BRANCH]', arr[1-1]);
	Result = Result.replace('[OFFICE]', arr[2-1]);
	Result = Result.replace('[KIND]', arr[3-1]);
	Result = Result.replace('[SUBKIND]', arr[4-1]);
	Result = Result.replace('[CURRENCY]', arr[5-1]);

	return Result;
}

function getCustomData(DB, type, vsp)
{
	var Result, objConn, objRecords, objFields, data, arr;

	Result = new Array();

	objConn = new ActiveXObject('ADODB.Connection');
	objConn.Open(getCustomSTR(DB));
	objRecords = objConn.Execute(getCustomSQL(type, vsp));
	if((!(objRecords.EOF)))
	{
		for((objRecords.MoveFirst()); (!(objRecords.EOF)); (objRecords.MoveNext()))
		{
			objFields = new Enumerator(objRecords.Fields);
			if((!(objFields.atEnd())))
			{
				data = new Array();
				data['KEY'] = new Array();
				data['KEY']['REST'] = new Array();
				data['KEY']['HEAD'] = new Array();
				data['KEY']['DATA'] = new Array();
				data['VAL'] = new Array();
				data['VAL']['REST'] = new Array();
				data['VAL']['HEAD'] = new Array();
				data['VAL']['DATA'] = new Array();
				for((objFields.moveFirst()); (!(objFields.atEnd())); (objFields.moveNext()))
				{
					arr = objFields.item().Name.split(':');
					switch(arr[0])
					{
						case 'KEY':
						{
							data['KEY']['HEAD'].push(arr[1]);
							data['KEY']['DATA'].push(objFields.item().Value);
						}
						break;
						case 'VAL':
						{
							data['VAL']['HEAD'].push(arr[1]);
							data['VAL']['DATA'].push(objFields.item().Value);
						}
						break;
						default:
						{
							throw new Error('Unsupported argument passed.');
						}
						break;
					}
				}
				data['KEY']['REST'].push(data['KEY']['HEAD'].join(';'));
				data['KEY']['REST'].push(data['KEY']['DATA'].join(';'));
				data['VAL']['REST'].push(data['VAL']['HEAD'].join(';'));
				data['VAL']['REST'].push(data['VAL']['DATA'].join(';'));
				Result[data['KEY']['REST'].join(':')] = data['VAL']['REST'].join(':')
			}
		}
	}
	objConn.Close();

	return Result;
}

function getCustomDiffs(type, vsp)
{
	var Result, DB, flag, data, arr, tmp, flt, i, k, m;

	Result = new Array();

	DB = new Array('FRONT', 'BACK');

	data = new Array();
	for(i in DB)
	{
		arr = getCustomData(DB[i], type, vsp);
		for(k in arr)
		{
			if((!(data[k])))
			{
				data[k] = new Array();
			}
			data[k][i] = arr[k];
		}
	}

	arr = new Array();
	for(k in data)
	{
		for(m in data[k])
		{
			for(i in DB)
			{
				if((!(data[k][i])))
				{
					flt = new Array();
					tmp = data[k][m].split(':');
					flt.push(tmp[0]);
					tmp = new String();
					flt.push(tmp);
					data[k][i] = flt.join(':');
				}
			}
		}
		for(i in DB)
		{
			flag = 0;
			for(m in data[k])
			{
				if((!(data[k][m] == data[k][i])))
				{
					flag = 1;
				}
			}
			if((!(flag == 0)))
			{
				if((!(arr[k])))
				{
					arr[k] = new Array();
				}
				arr[k][i] = data[k][i];
			}
		}
	}

	for(k in arr)
	{
		for(i in arr[k])
		{
			tmp = arr[k][i].split(':');
			for(m in tmp)
			{
				flt = new Array();
				flt = tmp[m].split(';');
				if(((m % 2) == 0))
				{
					flt.reverse();
					flt.push('');
					flt.reverse();
				}
				else
				{
					flt.reverse();
					flt.push(DB[i]);
					flt.reverse();
				}
				tmp[m] = flt.join(';');
			}
			arr[k][i] = tmp.join(':');
		}
		tmp = k.split(':');
		for(m in tmp)
		{
			flt = new Array();
			flt = tmp[m].split(';');
			if(((m % 2) == 0))
			{
				flt.reverse();
				flt.push('');
				flt.reverse();
			}
			else
			{
				flt.reverse();
				flt.push(type);
				flt.reverse();
			}
			tmp[m] = flt.join(';');
		}
		Result[tmp.join(':')] = arr[k].join(':');
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
	for(i in types)
	{
		diffs = getCustomDiffs(types[i], vsp);
		for(k in diffs)
		{
			Result[k] = diffs[k];
		}
	}

	return Result;
}

function getVSPList(diffs)
{
	var Result, arr, i;

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
	content.push('  f915.');
	content.push('</title>');
	content.push('</head>');
	content.push('<body>');
	for(i in diffs)
	{
		values = i.split(':');
		content.push('<table>');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=0 cellspacing=0>');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=3 cellspacing=0>');
		for(k in values)
		{
			if((!((k % 2) == 0)) || ((k == 0)))
			{
				items = values[k].split(';');
				content.push('<tr>');
				for(m in items)
				{
					content.push('<th>');
					content.push(items[m]);
					content.push('</th>');
				}
				content.push('</tr>');
			}
		}
		content.push('</table>');
		content.push('</td>');
		content.push('</tr>');
		values = diffs[i].split(':');
		content.push('<tr>');
		content.push('<td>');
		content.push('<table border=1 cellpadding=3 cellspacing=0>');
		for(k in values)
		{
			if((!((k % 2) == 0)) || ((k == 0)))
			{
				items = values[k].split(';');
				content.push('<tr>');
				for(m in items)
				{
					content.push('<th>');
					content.push(items[m]);
					content.push('</th>');
				}
				content.push('</tr>');
			}
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
