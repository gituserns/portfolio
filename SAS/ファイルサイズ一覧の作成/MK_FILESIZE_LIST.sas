%macro MK_FILESIZE_LIST(dir=, lib=, in=, excel=);
%put --------------------------------------------------;
%put �t�@�C���T�C�Y�ꗗ�̍쐬;
%put MK_FILESIZE_LIST;
%put &=dir;		/*�Ώۃt�H���_�p�X���w��i�����̊g���q.sas7bdat��ΏۂƂ���A�T�u�t�H���_�͊܂܂Ȃ��j*/
%put &=lib;		/*���C�u�����Q�Ɩ����w��i����dir���w�肵�Ȃ��ꍇ�͕K�{�j*/
%put &=in;		/*���̓f�[�^���w��i�ȗ��A�X�y�[�X��؂�ŕ����w��A�ȗ������ꍇ�͑S�f�[�^�Z�b�g�Ώہj*/
%put &=excel;	/*�o�̓G�N�Z���̃t���p�X���w��*/
%put --------------------------------------------------;

	/*�t�H�[�}�b�g�G���[��h��*/
	options nofmterr;

	/*���[�J���}�N���ϐ��̒�`*/
	%local i dsnum inds in_comp_obs compress modate;

	/*�f�[�^�Z�b�g�������[�J���}�N���ϐ��֊i�[����*/
	%if %length(&in.) > 0 %then %do;
		%let dsnum = %sysfunc(countw(&in.,, s));
		%do i = 1 %to &dsnum.;
			%local ds&i.;
			%let ds&i. = %scan(&in., &i.,, s);
		%end;
	%end;
	%else %do;
		
		%if %length(%superq(dir)) > 0 %then %do;
			filename dir "%superq(dir)";
			data _dsname_list;
				attrib name length=$1000.;
				keep name;
				open_dir = dopen("dir") ;
				do i = 1 to dnum(open_dir);
					name = dread(open_dir, i);
					output;
				end;
				rc = dclose(open_dir);
			run;
			filename dir clear;
		%end;
		%else %do;
			ods output members = _dsname_list;
			proc datasets lib=&lib. memtype=(data);
			quit;
			ods output close;
		%end;
		
		/*�s�v�����ȃt�@�C���͏���*/
		data _null_;
			set _dsname_list;
			where
					index(lowcase(name), "�R�s�[") = 0
				and index(lowcase(name), "_bk.") = 0
				and index(lowcase(name), "_old.") = 0
				and index(lowcase(name), "_chk.") = 0
				%if %length(%superq(dir)) > 0 %then %do;
					and reverse(substr(strip(reverse(lowcase(name))), 1, 9)) = ".sas7bdat"
				%end;
			;
			dsnum + 1;
			call symputx("dsnum", dsnum);
			call symputx(cats("ds", dsnum), upcase(tranwrd(lowcase(name), ".sas7bdat", "")), "L");
		run;
	%end;
	%put --------------------------------------------------;
	%put �Ώۃf�[�^�Z�b�g�̃��[�J���}�N���ϐ����̃`�F�b�N;
	%put &=dsnum;
	%put --------------------------------------------------;
	%do i = 1 %to &dsnum.;
		%put ds&i. = &&ds&i..;
	%end;
	%put --------------------------------------------------;
	%put;

	/*���C�u�����Q�Ɩ��̐ݒ�*/
	%if %length(%superq(dir)) > 0 %then %do;
		libname lib "%superq(dir)" access=readonly;
	%end;
	%else %do;
		libname lib (&lib.) access=readonly;
	%end; 

	/*���̓f�[�^�Z�b�g���ƂɃ��[�v����*/
	%do i = 1 %to &dsnum.;

		/*�f�[�^�Z�b�g�̊i�[*/
		%let inds = &&ds&i..;

		%put;
		%put --------------------------------------------------;
		%put &i. / &dsnum. �Ԗڂ̃f�[�^�Z�b�g;
		%put &inds.;
		%put --------------------------------------------------;
		%put;

		/*�t�@�C���T�C�Y�̎擾�i�i���k����j�j*/
		ods output	EngineHost = _in_comp_enginehost(keep=Label1 nValue1 rename=(Label1=Label1_comp nValue1=nValue1_comp))
					Attributes = _in_comp_attributes(keep=Label2 nValue2);
		proc contents data = lib.&inds.
					  out  = _in_comp_cont_&i.
					  ;
		run;
		ods output close;

		/*���݂̃f�[�^�����k�ς݂��ۂ��`�F�b�N����*/
		data _null_;
			set _in_comp_cont_&i.(obs=1);
			call symputx("compress",	compress);
			call symputx("in_comp_obs",	nobs);
			call symputx("modate",		modate);
		run;

		/*���k���������Ă��Ȃ��ꍇ�A���k���ēǂݍ���*/
		/*���k�ς݂ł���΁A�f�[�^��ǂݍ��ޕK�v���Ȃ��A�����������Ȃ�*/
		%if %upcase(&compress.) ne CHAR %then %do;
			data _in_comp(compress = yes);
				set lib.&inds.;
			run;
			ods output	EngineHost = _in_comp_enginehost(keep=Label1 nValue1 rename=(Label1=Label1_comp nValue1=nValue1_comp))
						Attributes = _in_comp_attributes(keep=Label2 nValue2);
			proc contents data = _in_comp
						  out  = _in_comp_cont_&i.
						  ;
			run;
			ods output close;

			proc datasets lib=work noprint;
				delete _in_comp;
			quit;
		%end;

		/*�ϐ��̕��я�����̓f�[�^�Ɠ����ɂ���*/
		proc sort data = _in_comp_cont_&i.;
			by varnum;
		run;
		data _in_comp_cont_&i.;
			set _in_comp_cont_&i.;
			attrib
				format_length	length=$100. label="�o�͌`��"
				informat_length	length=$100. label="���͌`��"
			;
			if missing(format) then		format_length = "";
			else						format_length = cats(format, formatl, ".");
			if missing(informat) then	informat_length = "";
			else						informat_length = cats(informat, informl, ".");
		run;

		/*�f�[�^�̓Ǎ��݁i���k�Ȃ��j*/
		/*
			���k�Ȃ��Ńf�[�^��ǂݍ��񂾎��_�ŁA�f�B�X�N�s���G���[�ɂȂ邱�Ƃ�������邽�߁A
			1000�I�u�U�x�[�V�������́A�t�@�C���T�C�Y�̐���l��ێ�I�ɎZ�o��
			1000�I�u�U�x�[�V�����ȉ��͎��ۂ̃t�@�C���T�C�Y���擾����B

			�t�@�C���T�C�Y�i����l�j = �I�u�U�x�[�V������ �~ �I�u�U�x�[�V�����̃o�b�t�@�� �~ 1.02
			�o���I�ɁA���ۂ̃t�@�C���T�C�Y��1.02���悶���ꍇ�̐���l�𒴂���\���͋ɂ߂ĒႢ�B
		*/
		%if &in_comp_obs. > 1000 %then %do;
			data _in_nocomp;
				set lib.&inds.(obs=0);
			run;
		%end;
		%else %do;
			data _in_nocomp(compress = no);
				set lib.&inds.;
			run;
		%end;

		/*�t�@�C���T�C�Y�̎擾�i�i���k�Ȃ��j�j*/
		ods output	EngineHost = _in_nocomp_enginehost(keep=Label1 nValue1);
		proc contents data = _in_nocomp;
		run;
		ods output close;

		/*�t�@�C���T�C�Y�ꗗ�̍쐬*/
		data _datasize_list_temp;
			attrib
				id		length=8.
				label	length=$200.
				value	length=8.
			;
			keep
				id
				label
				value
			;
			set
				_in_comp_attributes
				_in_nocomp_enginehost
				_in_comp_enginehost
			;

			if strip(Label2) in ("�I�u�U�x�[�V������", "�ϐ��̐�", "�I�u�U�x�[�V�����̃o�b�t�@��") then do;
				id + 1;
				label = strip(Label2);
				value = nValue2;
				output;
			end;

			if strip(Label1) in ("�t�@�C���T�C�Y (�o�C�g)") then do;
				id + 1;
				label = "�t�@�C���T�C�YGB�i���k�Ȃ��E����l�j";
				value = nValue1 / 1000000000;
				output;
			end;

			if strip(Label1_comp) in ("�t�@�C���T�C�Y (�o�C�g)") then do;
				id + 1;
				label = "�t�@�C���T�C�YGB�i���k����j";
				value = nValue1_comp / 1000000000;
				output;
			end; 

		run;

		/*�]�u*/
		proc transpose data=_datasize_list_temp out=_datasize_list_tran(drop=_name_) prefix=var_;
			id id;
			idlabel label;
			var value;
		run;

		/*���X�g*/
		data _datasize_list_&i.;
			attrib
				dir		length=$100.	label="���C�u������"
				dsname	length=$32.		label="�f�[�^�Z�b�g��"
				modate	length=8.		label="�X�V��"	format=datetime20.
			;
			set _datasize_list_tran;

			%if %length(%superq(dir)) > 0 %then %do;
				dir	= strip(reverse(scan(reverse("%superq(dir)"), 1, "\")));
			%end;
			%else %do;
				dir	= strip("&lib.");
			%end;
			dsname	= strip("&inds.");
			modate  = &modate.;

			/*1000�I�u�U�x�[�V�������̏ꍇ�͐���l���Z�o����*/
			%if &in_comp_obs. > 1000 %then %do;
				var_4 = var_1 * var_3 * 1.02 / 1000000000;
			%end;

			format var_4 var_5 best.;
		run;

		/*�s�v�f�[�^�̍폜*/
		proc datasets lib=work noprint;
			delete
				_in_nocomp
				_in_comp_attributes
				_in_nocomp_enginehost
				_in_comp_enginehost
				_datasize_list_temp
				_datasize_list_tran
			;
		quit;

	%end;

	/*�W��*/
	data _datasize_list_fin;
		set _datasize_list_1 - _datasize_list_&dsnum.;
	run;


	/*�G�N�Z���o��*/
	ods excel file="%superq(excel)" 	options(
											sheet_name		="�t�@�C���T�C�Y�ꗗ"	/*�V�[�g��*/
											sheet_interval	="proc"					/*PROC���ƂɃV�[�g�𕪂���*/
											start_at		="2,2"					/*�J�n�s��*/
											autofilter		="all"					/*�I�[�g�t�B���^*/
											embedded_titles	="off"					/*�^�C�g���Ȃ�*/
										);
	proc print data=_datasize_list_fin label noobs;
	run;

	%do i = 1 %to &dsnum.;
		ods excel options(
			sheet_name		="&&ds&i.."
			sheet_interval	="proc"	
			start_at		="2,2"	
			autofilter		="all"	
			embedded_titles	="off"
		);
		proc print data=_in_comp_cont_&i. label;
			var
				varnum
				name
				label
				type
				length
				format_length
				informat_length
			;
		run;
	%end;

	ods excel close;

	/*�s�v�f�[�^�̍폜*/
	proc datasets lib=work noprint;
		delete
			_datasize_list_:
		;
	quit;

	libname lib clear;

	options fmterr;

%mend MK_FILESIZE_LIST;


/*�W�J��1:�t�H���_�����̑SSAS�f�[�^�Z�b�g�ꗗ�̎擾*/
/*
%MK_FILESIZE_LIST(
	dir		= C:\library\FCS,
	excel	= C:\Users\�t�@�C���T�C�Y�ꗗ\�t�@�C���T�C�Y�ꗗ_FCS.xlsx
);
*/

/*�W�J��2:�t�H���_�����̎w�肵��SAS�f�[�^�Z�b�g�ꗗ�̎擾*/
/*
%MK_FILESIZE_LIST(
	dir		= C:\librar\DM,
	in		=
			D_0032_01
			D_0023_01
			D_0024_01
			D_0025_01
			D_0026_01
			D_0027_01
			D_0028_01
			D_0029_01
			D_0065_01
			D_0064_01
			D_0031_01
			D_0046_01
			,
	excel	=  C:\Users\�t�@�C���T�C�Y�ꗗ\�t�@�C���T�C�Y�ꗗ_DM.xlsx
);
*/

/*�W�J��3:���C�u�����Q�Ɩ�������SAS�f�[�^�Z�b�g�ꗗ�̎擾*/
/*
%MK_FILESIZE_LIST(
	lib		= togo,
	excel	= C:\Users\�t�@�C���T�C�Y�ꗗ\�t�@�C���T�C�Y�ꗗ_�����}�X�^.xlsx
);
*/
