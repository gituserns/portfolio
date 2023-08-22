%macro mathematical_sets(lib1=, in1=, lib2=, in2=, key=, nodupkey=Y, path=, name=);
%put --------------------------------------------------;
%put  mathematical_sets;
%put 2つのデータ間の集合と結合率を算出する;
%put &=lib1;	/*データAのライブラリ参照名が割り当てられていない場合指定する（省略可）*/
%put &=in1;		/*1つ目のデータAを指定する（データセットオプションの指定可）*/
%put &=lib2;	/*データBのライブラリ参照名が割り当てられていない場合指定する（省略可）*/
%put &=in2;		/*2つ目のデータBを指定する（データセットオプションの指定可）*/
%put &=key;		/*要素となるキーを指定する（事前に2つのデータの変数属性を揃えておくこと）*/
%put &=nodupkey;/*キーによる重複削除をする場合はYを指定する（既定値=Y）*/
%put &=path;	/*ファイルの出力先を指定（末尾のバックスラッシュ"\"は不要）*/
%put &=name;	/*ファイル名を指定(ファイル名.html、ファイル名.xlsxが作成される)*/
%put --------------------------------------------------;

	/*フォーマットエラーを防ぐ*/
	options nofmterr;

	/*ローカルマクロ変数の定義*/
	%local n n1 n2 n12;

	/*キーで重複削除し、データを読込む*/
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

	/*集合のカウント*/
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

		/*和集合 n(A ∪ B)*/
		n = 1;

		/*集合A n(A)*/
		if in1 then n1 = 1;
		else n1 = .;

		/*集合B n(B)*/
		if in2 then n2 = 1;
		else n2 = .;

		/*共通集合 n(A ∩ B)*/
		if in1 and in2 then n12 = 1;
		else n12 = .;

		output _mathematical_sets_cnt;
		
		/*エクセル用追加*/
		if in1 and ^in2 then output _mathematical_sets_xls1;
		if ^in1 and in2 then output _mathematical_sets_xls2;
		if in1 and in2  then output _mathematical_sets_xls12;
	run;

	/*集計*/
	proc summary data = _mathematical_sets_cnt nway;
		var n n1 n2 n12;
		output
			out = _mathematical_sets_sum
			n=
		;
	run;

	/*各集合をマクロ変数化する*/
	data _null_;
		set _mathematical_sets_sum;
		call symputx("n", n);
		call symputx("n1", n1);
		call symputx("n2", n2);
		call symputx("n12", n12);
	run;

	/*ヘッダー*/
	data _mathematical_sets_header;
		attrib
			taisyo	length=$200.	label="対象"
			naiyo	length=$200.	label="内容"
		;
		taisyo	="データA";
		naiyo	="%superq(in1)";
		output;

		taisyo	="データB";
		naiyo	="%superq(in2)";
		output;

		taisyo	="キー（要素）";
		naiyo	="%superq(key)";
		output;

		taisyo	="集計前のキーによる重複削除";
		%if %upcase(&nodupkey.) = Y %then %do;
			naiyo="あり";
		%end;
		%else %do;
			naiyo="なし";
		%end;
		output;
	run;

	/*集計の整理*/
	data _mathematical_sets_fin;
		attrib
			naiyo	length=$50.	label="内容"
			sets	length=$50.	label="集合"
			type	length=$20.	label="タイプ"
			value	length=$20.	label="値"
		;
		
		naiyo	="和集合";
		sets	= "n (A ∪ B)";
		type	="件数";
		value	=put(&n., comma.);
		output;

		naiyo	="集合A";
		sets	= "n (A)";
		type	="件数";
		value	=put(&n1., comma.);
		output;

		naiyo	="集合B";
		sets	= "n (B)";
		type	="件数";
		value	=put(&n2., comma.);
		output;

		naiyo	="共通集合";
		sets	= "n (A ∩ B)";
		type	="件数";
		value	=put(&n12., comma.);
		output;

		naiyo	="集合Aのみに含まれる";
		sets	= "n (A ∩ ^B)";
		type	="件数";
		value	=put(&n1.-&n12., comma.);
		output;

		naiyo	="集合Bのみに含まれる";
		sets	= "n (^A ∩ B)";
		type	="件数";
		value	=put(&n2.-&n12., comma.);
		output;

		naiyo	="和集合のうち共通集合の割合";
		sets	= "n (A ∩ B) / n (A ∪ B)";
		type	="割合（％）";
		value	=put(&n12. / &n. * 100, 8.2);
		output;

		naiyo	="和集合のうち集合Aの割合";
		sets	= "n (A) / n (A ∪ B)";
		type	="割合（％）";
		value	=put(&n1. / &n. * 100, 8.2);
		output;

		naiyo	="和集合のうち集合Bの割合";
		sets	= "n (B) / n (A ∪ B)";
		type	="割合（％）";
		value	=put(&n2. / &n. * 100, 8.2);
		output;

		naiyo	="集合Aのうち共通集合の割合";
		sets	= "n (A ∩ B) / n (A)";
		type	="割合（％）";
		value	=put(&n12. / &n1. * 100, 8.2);
		output;

		naiyo	="集合Bのうち共通集合の割合";
		sets	= "n (A ∩ B) / n (B)";
		type	="割合（％）";
		value	=put(&n12. / &n2. * 100, 8.2);
		output;

	run;

	/*HTML出力*/
	ods html path = "%superq(path)" file ="%superq(name).html";
	proc print data = _mathematical_sets_header label noobs;
	run;
	proc print data = _mathematical_sets_fin label noobs;
	run;
	ods html close;

	/*エクセル出力*/
	ods excel file="%superq(path)\%superq(name).xlsx" 	options(
											sheet_name		="集合一覧"	/*シート名*/
											sheet_interval	="none"		/*PROCごとにシートを分ける*/
											start_at		="2,2"		/*開始行列*/
											autofilter		="all"		/*オートフィルタ*/
											embedded_titles	="off"		/*タイトルなし*/
										);
	proc print data = _mathematical_sets_header label noobs;
	run;
	proc print data = _mathematical_sets_fin label noobs;
	run;

	ods excel options(
		sheet_name		="データAのみに存在"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_mathematical_sets_xls1 label;
		var &key.;
	run;

	ods excel options(
		sheet_name		="データBのみに存在"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_mathematical_sets_xls2 label;
		var &key.;
	run;

	ods excel options(
		sheet_name		="データAとB両方に存在"
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


/*使用例*/
/*
%mathematical_sets(
	lib1	=C:\library, 
	in1		=d_0030_01, 
	lib2	=, 
	in2		=T_OWNER_TP_SAS_RNK, 
	key		=TEN_CD,
	nodupkey=Y,
	path	=C:\Users\結合率, 
	name	= 01_結合率（店舗コード）
);

%mathematical_sets(
	lib1	=C:\library, 
	in1		=d_0030_01, 
	lib2	=, 
	in2		=T_OWNER_TP_SAS_RNK, 
	key		=TEN_CD OWNER_CD,
	nodupkey=Y,
	path	=C:\Users\結合率, 
	name	= 02_結合率（店舗コード・オーナーコード）
);
*/
