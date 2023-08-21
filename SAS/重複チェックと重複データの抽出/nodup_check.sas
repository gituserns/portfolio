%macro nodup_check(lib=, in=, key=_all_,  path=, name=);
%put --------------------------------------------------;
%put nodup_check;
%put �d���`�F�b�N�Əd���f�[�^�̒��o;
%put &=lib;		/*���C�u�����Q�Ɛ�̎w��i�ȗ��j*/
%put &=in;		/*�`�F�b�N�Ώۃf�[�^�Z�b�g���w��i�f�[�^�Z�b�g�I�v�V�����̎w��j*/
%put &=key;		/*�`�F�b�N�Ώ�KEY���w��i����l=_all_:�S�ϐ��j*/
%put &=path;	/*�t�@�C���̏o�͐���w��i�����̃o�b�N�X���b�V��"\"�͕s�v�j*/
%put &=name;	/*�t�@�C�������w��(�t�@�C����.html�A�t�@�C����.xlsx���쐬�����)*/
%put --------------------------------------------------;

	/*�t�H�[�}�b�g�G���[��h��*/
	options nofmterr compress=yes;

	/*���[�J���}�N���ϐ��̒�`*/
	%local i num_indata num_dupdata;

	/*���C�u�����Q�Ɩ��̐ݒ�ƃf�[�^�̓ǂݍ���*/
	%if %length(%superq(lib)) > 0 %then %do;
		libname lib &lib. access=readonly;
		proc sort data=lib.&in. out=_nodup_check_data;
			by &key.;
		run;
		libname lib clear;
	%end;
	%else %do;
		proc sort data=&in. out=_nodup_check_data;
			by &key.;
		run;
	%end;

	/*���̓f�[�^�̃��R�[�h���̎擾*/
	%let num_indata = &sysnobs.;
	%if &num_indata. = 0 %then %put WARNING:���̓f�[�^�Z�b�g�͋�ł��B;

	/*�ϐ����̎擾*/
	data 
		_nodup_check_cont_key(keep=&key.)
		_nodup_check_cont_all
		;
		format &key.;
		set _nodup_check_data;
		stop;
	run;
	proc contents data=_nodup_check_cont_key out=_nodup_check_cont_key2 noprint;
	run;
	proc contents data=_nodup_check_cont_all out=_nodup_check_cont_all2 noprint;
	run;
	proc sort data=_nodup_check_cont_key2;
		by varnum;
	run;
	proc sort data=_nodup_check_cont_all2;
		by varnum;
	run;	

	/*�ϐ��̃��[�J���}�N���ϐ���*/
	data _null_;
		set _nodup_check_cont_key2;
		call symputx("num_key", _n_, "L");
		call symputx(cats("key", _n_), name, "L");
	run;
	data _null_;
		set _nodup_check_cont_all2;
		call symputx("num_var", _n_, "L");
		call symputx(cats("var", _n_), name,	"L");
		call symputx(cats("type", _n_),type,	"L");
		call symputx(cats("len", _n_), length,	"L");
	run;

	/*�d���f�[�^�̏o��*/
	proc sort data=_nodup_check_data(keep=&key.) 
									out=_nodup_check_nodup 
									dupout=_nodup_check_dup
									nodupkey;
		by &key.;
	run;
	
	/*�d���f�[�^�̏d���폜*/
	proc sort data=_nodup_check_dup nodupkey;
		by &key.;
	run;

	/*_nodup_check_dup��OBS���ɂ���������*/
	%if &sysnobs. > 0 %then %do;
		data 
			_nodup_check_data_all(drop=_sabun_var)
			_nodup_check_sabun(keep=_dup_no _sabun_var)
			;

			drop 
				_logout_cnt
				%do i = 1 %to &num_var.;
					_&&var&i..
					_sabun_cnt&i.
				%end;
			;

			attrib
				_dup_no 		length=8. 		label="�d��No."
				_dup_key		length=$10000.	label="�d���L�["
				_sabun_var		length=$10000.	label="�����̂���ϐ�"

				/*RETAIN�p�ϐ�*/
				%do i = 1 %to &num_var.;
					_&&var&i..
					%if &&type&i.. = 1 %then %do;
						length=8.
					%end;
					%else %do;
						length=$&&len&i..
					%end;
				%end;	
			;

			retain
				_sabun_var
				%do i = 1 %to &num_var.;
					_&&var&i..
				%end;
			;

			merge
				_nodup_check_data(in=in1)
				_nodup_check_dup(in=in2)
			;
			by &key.;
			if in1 and in2;

			/*�`�F�b�N�L�[�̎擾*/
			_dup_key = strip("&key.");

			/*�d�����R�[�h�̃��O�ւ̏o��*/
			_logout_cnt + 1;
			if _logout_cnt <= 10 then put "WARNING:" _n_ +(-1) "�I�u�U�x�[�V�����ڂ��d�����Ă��܂��B" +1 (&key.)(=);
			if _logout_cnt = 10 then put "WARNING:�d������10�ɒB�����̂ŁA���O�ւ̏o�͂��~�߂܂��B";
			
			
			/*�d��No.�̎擾*/
			if first.&&key&num_key.. then _dup_no + 1;

			/*�����̂���ϐ��̎擾*/
			if first.&&key&num_key.. then do;
				call missing(of _sabun_var);
				%do i = 1 %to &num_var.;
					_sabun_cnt&i. = 0;
					_&&var&i.. = &&var&i..;
				%end;
			end;
			%do i = 1 %to &num_var.;
				if _&&var&i.. ne &&var&i.. and _sabun_cnt&i. = 0 then do;
					_sabun_cnt&i. + 1;

					if strip(vlabel(&&var&i..)) ne strip(vname(&&var&i..)) then _sabun_var = catx(" ", _sabun_var, cats(vlabel(&&var&i..), ":", vname(&&var&i..)));
					else _sabun_var = catx(" ", _sabun_var, vname(&&var&i..));
				end;
			%end;

			/*�o��*/
			output _nodup_check_data_all;
			if last.&&key&num_key.. and not missing(_sabun_var) then output _nodup_check_sabun;

		run;

		/*�����̂���ϐ��̎擾*/
		data _nodup_check_data_all_fin;

			format
				_dup_no
				_dup_key
				_sabun_var
				&key.
			;
			merge
				_nodup_check_data_all(in=in1)
				_nodup_check_sabun(in=in2)
			;
			by _dup_no;
			if in1;
		run;

		/*�d���f�[�^�̃��R�[�h���̎擾*/
		%let num_dupdata = &sysnobs.;

		/*���z�̏W�v*/
		proc freq data = _nodup_check_data_all_fin noprint;
			table _sabun_var / nocol norow nopercent out = _nodup_check_sabun_var_freq;
		run;

		/*����*/
		data _nodup_check_sabun_var_freq_fin;
			attrib _dup_pattern	length=$200. label="�d���p�^�[��";
			set 
				_nodup_check_sabun_var_freq(in=in0 obs=1)
				_nodup_check_sabun_var_freq
				end = eof
			;
			if in0 then do;
				_dup_pattern = "�d���Ȃ�";
				_sabun_var = "";
				count = sum(&num_indata., -&num_dupdata.);
				percent = count / &num_indata. * 100;
			end;
			else do;
				_dup_pattern = cats("�d���p�^�[��", _n_ - 1);
				percent = count / &num_indata. * 100;
			end;

			output;

			/*���v�X�e�[�g�����g*/
			_sum_count + count;
			_sum_percent + percent;
			
			if eof then do;
				_dup_pattern = "���v";
				_sabun_var = "";
				if _sum_count ne &num_indata. then put "WARNING:���v�x�������̓f�[�^�̃��R�[�h���ƕs��v"(_sum_count)(=);
				count = _sum_count;
				if not(99.99 < _sum_percent < 100.01) then put "WARNING:���v��100�p�[�Z���g�ɂȂ��Ă��Ȃ�" (_sum_percent)(=);
				percent = _sum_percent;
				output;
			end;

			drop
				_sum_count
				_sum_percent
			;

			format count comma8.;
		run;

	%end;
	%else %do;
		%put NOTE:�d���f�[�^�͑��݂��܂���B;
		data _nodup_check_data_all_fin;
			attrib
				_dup_key	length=$10000.	label="�d���L�["
				_sabun_var	length=$20.		label="�R�����g"
			;

			/*�`�F�b�N�L�[�̎擾*/
			_dup_key = strip("&key.");

			/*�R�����g*/
			_sabun_var = "�d�����R�[�h�Ȃ�";
		run;

		data _nodup_check_sabun_var_freq_fin;
			attrib 
				_dup_pattern	length=$200.	label="�d���p�^�[��"
				count			length=8.		label="�x��"
				percent			length=8.		label="���v�x���̃p�[�Z���g"

			;
			_dup_pattern = "�d���Ȃ�";
			count = &num_indata.;
			percent = 100;
		run;

	%end;

	/*�d�������Ώ�*/
	data _nodup_check_list;
		attrib
			komoku	length=$100.	label="����"
			naiyo	length=$1000.	label="���e"
		;
		komoku	= "���C�u�����Q�Ɛ�";
		naiyo	= "%superq(lib)";
		output;

		komoku	= "�e�[�u����";
		naiyo	= "%superq(in)";
		output;

		komoku	= "�����L�[";
		naiyo	= "%superq(key)";
		output;
	run;

	/*HTML�o��*/
	ods html path = "%superq(path)" file ="%superq(name).html";
	title "�d������";
	proc print data = _nodup_check_list label noobs;
	run;
	title "�d���p�^�[�����z";
	proc print data = _nodup_check_sabun_var_freq_fin label noobs;
	run;
	ods html close;
	title "";
	/*�G�N�Z���o��*/
	ods excel file="%superq(path)\%superq(name).xlsx" 	options(
											sheet_name		="�d���p�^�[�����z"	/*�V�[�g��*/
											sheet_interval	="none"				/*PROC���ƂɃV�[�g�𕪂���*/
											start_at		="2,2"				/*�J�n�s��*/
											autofilter		="all"				/*�I�[�g�t�B���^*/
											embedded_titles	="off"				/*�^�C�g���Ȃ�*/
										);
	proc print data = _nodup_check_list label noobs;
	run;
	proc print data = _nodup_check_sabun_var_freq_fin label noobs;
	run;
/*
	ods excel options(
		sheet_name		="�d���f�[�^"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_nodup_check_data_all_fin label noobs;
	run;
*/
	ods excel close;

	/*�d���f�[�^����������ƃ������s�����N�������߁AEXPORT�v���V�W���ŏo�͂���*/
	proc export data=_nodup_check_data_all_fin
				outfile="%superq(path)\%superq(name).xlsx"
				dbms=xlsx
				label;
		sheet="�d���f�[�^"n;
	run;

	/*�s�v�f�[�^�̍폜*/
	proc datasets lib=work noprint;
		delete _nodup_check_:;
	quit;

	options fmterr compress=no;

%mend nodup_check;


/*�g�p��*/
/*
%nodup_check(
	lib 	="C:\library",
	in		=D_0023_01(keep=IM_TEN_CD IM_OWNER_CD IM_JIGYOSHA_NO IM_TAX_KBN),
	key		=TEN_CD,
	path	=C:\Users\�d���f�[�^,
	name	=�d������_D_0023_01
);

%nodup_check(
	lib	=,
	in	= D_0030_01_INVOICE,
	key	= TEN_CD KEIYAK_SDATE KEIYAK_EDATE,
	path= C:\Users\�d���f�[�^,
	name= �d������_D_0030_01_INVOICE
);

%nodup_check(
	lib	=,
	in	= D_0030_01_INVOICE_NODUP,
	key	= TEN_CD KEIYAK_SDATE KEIYAK_EDATE,
	path= C:\Users\�d���f�[�^,
	name= �d������_D_0030_01_INVOICE_NODUP
);
*/
