%macro mathematical_sets(lib1=, in1=, lib2=, in2=, key=, nodupkey=Y, path=, name=);
%put --------------------------------------------------;
%put  mathematical_sets;
%put 2�̃f�[�^�Ԃ̏W���ƌ��������Z�o����;
%put &=lib1;	/*�f�[�^A�̃��C�u�����Q�Ɩ������蓖�Ă��Ă��Ȃ��ꍇ�w�肷��i�ȗ��j*/
%put &=in1;		/*1�ڂ̃f�[�^A���w�肷��i�f�[�^�Z�b�g�I�v�V�����̎w��j*/
%put &=lib2;	/*�f�[�^B�̃��C�u�����Q�Ɩ������蓖�Ă��Ă��Ȃ��ꍇ�w�肷��i�ȗ��j*/
%put &=in2;		/*2�ڂ̃f�[�^B���w�肷��i�f�[�^�Z�b�g�I�v�V�����̎w��j*/
%put &=key;		/*�v�f�ƂȂ�L�[���w�肷��i���O��2�̃f�[�^�̕ϐ������𑵂��Ă������Ɓj*/
%put &=nodupkey;/*�L�[�ɂ��d���폜������ꍇ��Y���w�肷��i����l=Y�j*/
%put &=path;	/*�t�@�C���̏o�͐���w��i�����̃o�b�N�X���b�V��"\"�͕s�v�j*/
%put &=name;	/*�t�@�C�������w��(�t�@�C����.html�A�t�@�C����.xlsx���쐬�����)*/
%put --------------------------------------------------;

	/*�t�H�[�}�b�g�G���[��h��*/
	options nofmterr;

	/*���[�J���}�N���ϐ��̒�`*/
	%local n n1 n2 n12;

	/*�L�[�ŏd���폜���A�f�[�^��Ǎ���*/
	%if %length(%superq(lib1)) > 0 %then %do;
		libname lib1 "%superq(lib1)" access = readonly;
		proc sort data = lib1.&in1. out = _mathematical_sets_in1
			%if %upcase(&nodupkey.) = Y %then nodupkey;
			;
			by &key.;
		run;
		libname lib1 clear;
	%end;
	%else %do;
		proc sort data=&in1. out=_mathematical_sets_in1
			%if %upcase(&nodupkey.) = Y %then nodupkey;
			;
			by &key.;
		run;		
	%end;

	%if %length(%superq(lib2)) > 0 %then %do;
		libname lib2 "%superq(lib2)" access = readonly;
		proc sort data = lib2.&in2. out = _mathematical_sets_in2
			%if %upcase(&nodupkey.) = Y %then nodupkey;
			;
			by &key.;
		run;
		libname lib2 clear;
	%end;
	%else %do;
		proc sort data=&in2. out=_mathematical_sets_in2
			%if %upcase(&nodupkey.) = Y %then nodupkey;
			;
			by &key.;
		run;		
	%end;

	/*�W���̃J�E���g*/
	data 
		_mathematical_sets_cnt
		_mathematical_sets_xls1(keep=&key.)
		_mathematical_sets_xls2(keep=&key.)
		_mathematical_sets_xls12(keep=&key.)
		;
		merge
			_mathematical_sets_in1(in = in1)
			_mathematical_sets_in2(in = in2)
		;
		by &key.;

		/*�a�W�� n(A �� B)*/
		n = 1;

		/*�W��A n(A)*/
		if in1 then n1 = 1;
		else n1 = .;

		/*�W��B n(B)*/
		if in2 then n2 = 1;
		else n2 = .;

		/*���ʏW�� n(A �� B)*/
		if in1 and in2 then n12 = 1;
		else n12 = .;

		output _mathematical_sets_cnt;
		
		/*�G�N�Z���p�ǉ�*/
		if in1 and ^in2 then output _mathematical_sets_xls1;
		if ^in1 and in2 then output _mathematical_sets_xls2;
		if in1 and in2  then output _mathematical_sets_xls12;
	run;

	/*�W�v*/
	proc summary data = _mathematical_sets_cnt nway;
		var n n1 n2 n12;
		output
			out = _mathematical_sets_sum
			n=
		;
	run;

	/*�e�W�����}�N���ϐ�������*/
	data _null_;
		set _mathematical_sets_sum;
		call symputx("n", n);
		call symputx("n1", n1);
		call symputx("n2", n2);
		call symputx("n12", n12);
	run;

	/*�w�b�_�[*/
	data _mathematical_sets_header;
		attrib
			taisyo	length=$200.	label="�Ώ�"
			naiyo	length=$200.	label="���e"
		;
		taisyo	="�f�[�^A";
		naiyo	="%superq(in1)";
		output;

		taisyo	="�f�[�^B";
		naiyo	="%superq(in2)";
		output;

		taisyo	="�L�[�i�v�f�j";
		naiyo	="%superq(key)";
		output;

		taisyo	="�W�v�O�̃L�[�ɂ��d���폜";
		%if %upcase(&nodupkey.) = Y %then %do;
			naiyo="����";
		%end;
		%else %do;
			naiyo="�Ȃ�";
		%end;
		output;
	run;

	/*�W�v�̐���*/
	data _mathematical_sets_fin;
		attrib
			naiyo	length=$50.	label="���e"
			sets	length=$50.	label="�W��"
			type	length=$20.	label="�^�C�v"
			value	length=$20.	label="�l"
		;
		
		naiyo	="�a�W��";
		sets	= "n (A �� B)";
		type	="����";
		value	=put(&n., comma.);
		output;

		naiyo	="�W��A";
		sets	= "n (A)";
		type	="����";
		value	=put(&n1., comma.);
		output;

		naiyo	="�W��B";
		sets	= "n (B)";
		type	="����";
		value	=put(&n2., comma.);
		output;

		naiyo	="���ʏW��";
		sets	= "n (A �� B)";
		type	="����";
		value	=put(&n12., comma.);
		output;

		naiyo	="�W��A�݂̂Ɋ܂܂��";
		sets	= "n (A �� ^B)";
		type	="����";
		value	=put(&n1.-&n12., comma.);
		output;

		naiyo	="�W��B�݂̂Ɋ܂܂��";
		sets	= "n (^A �� B)";
		type	="����";
		value	=put(&n2.-&n12., comma.);
		output;

		naiyo	="�a�W���̂������ʏW���̊���";
		sets	= "n (A �� B) / n (A �� B)";
		type	="�����i���j";
		value	=put(&n12. / &n. * 100, 8.2);
		output;

		naiyo	="�a�W���̂����W��A�̊���";
		sets	= "n (A) / n (A �� B)";
		type	="�����i���j";
		value	=put(&n1. / &n. * 100, 8.2);
		output;

		naiyo	="�a�W���̂����W��B�̊���";
		sets	= "n (B) / n (A �� B)";
		type	="�����i���j";
		value	=put(&n2. / &n. * 100, 8.2);
		output;

		naiyo	="�W��A�̂������ʏW���̊���";
		sets	= "n (A �� B) / n (A)";
		type	="�����i���j";
		value	=put(&n12. / &n1. * 100, 8.2);
		output;

		naiyo	="�W��B�̂������ʏW���̊���";
		sets	= "n (A �� B) / n (B)";
		type	="�����i���j";
		value	=put(&n12. / &n2. * 100, 8.2);
		output;

	run;

	/*HTML�o��*/
	ods html path = "%superq(path)" file ="%superq(name).html";
	proc print data = _mathematical_sets_header label noobs;
	run;
	proc print data = _mathematical_sets_fin label noobs;
	run;
	ods html close;

	/*�G�N�Z���o��*/
	ods excel file="%superq(path)\%superq(name).xlsx" 	options(
											sheet_name		="�W���ꗗ"	/*�V�[�g��*/
											sheet_interval	="none"		/*PROC���ƂɃV�[�g�𕪂���*/
											start_at		="2,2"		/*�J�n�s��*/
											autofilter		="all"		/*�I�[�g�t�B���^*/
											embedded_titles	="off"		/*�^�C�g���Ȃ�*/
										);
	proc print data = _mathematical_sets_header label noobs;
	run;
	proc print data = _mathematical_sets_fin label noobs;
	run;

	ods excel options(
		sheet_name		="�f�[�^A�݂̂ɑ���"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_mathematical_sets_xls1 label;
		var &key.;
	run;

	ods excel options(
		sheet_name		="�f�[�^B�݂̂ɑ���"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_mathematical_sets_xls2 label;
		var &key.;
	run;

	ods excel options(
		sheet_name		="�f�[�^A��B�����ɑ���"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_mathematical_sets_xls12 label;
		var &key.;
	run;

	ods excel close;

	options fmterr;

%mend mathematical_sets;


/*�g�p��*/
/*
%mathematical_sets(
	lib1	=C:\library, 
	in1		=d_0030_01, 
	lib2	=, 
	in2		=T_OWNER_TP_SAS_RNK, 
	key		=TEN_CD,
	nodupkey=Y,
	path	=C:\Users\������, 
	name	= 01_�������i�X�܃R�[�h�j
);

%mathematical_sets(
	lib1	=C:\library, 
	in1		=d_0030_01, 
	lib2	=, 
	in2		=T_OWNER_TP_SAS_RNK, 
	key		=TEN_CD OWNER_CD,
	nodupkey=Y,
	path	=C:\Users\������, 
	name	= 02_�������i�X�܃R�[�h�E�I�[�i�[�R�[�h�j
);
*/
