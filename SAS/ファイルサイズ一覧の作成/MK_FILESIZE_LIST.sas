%macro MK_FILESIZE_LIST(dir=, lib=, in=, excel=);
%put --------------------------------------------------;
%put ファイルサイズ一覧の作成;
%put MK_FILESIZE_LIST;
%put &=dir;		/*対象フォルダパスを指定（直下の拡張子.sas7bdatを対象とする、サブフォルダは含まない）*/
%put &=lib;		/*ライブラリ参照名を指定（引数dirを指定しない場合は必須）*/
%put &=in;		/*入力データを指定（省略可、スペース区切りで複数指定可、省略した場合は全データセット対象）*/
%put &=excel;	/*出力エクセルのフルパスを指定*/
%put --------------------------------------------------;

	/*フォーマットエラーを防ぐ*/
	options nofmterr;

	/*ローカルマクロ変数の定義*/
	%local i dsnum inds in_comp_obs compress modate;

	/*データセット名をローカルマクロ変数へ格納する*/
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
		
		/*不要そうなファイルは除く*/
		data _null_;
			set _dsname_list;
			where
					index(lowcase(name), "コピー") = 0
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
	%put 対象データセットのローカルマクロ変数化のチェック;
	%put &=dsnum;
	%put --------------------------------------------------;
	%do i = 1 %to &dsnum.;
		%put ds&i. = &&ds&i..;
	%end;
	%put --------------------------------------------------;
	%put;

	/*ライブラリ参照名の設定*/
	%if %length(%superq(dir)) > 0 %then %do;
		libname lib "%superq(dir)" access=readonly;
	%end;
	%else %do;
		libname lib (&lib.) access=readonly;
	%end; 

	/*入力データセットごとにループ処理*/
	%do i = 1 %to &dsnum.;

		/*データセットの格納*/
		%let inds = &&ds&i..;

		%put;
		%put --------------------------------------------------;
		%put &i. / &dsnum. 番目のデータセット;
		%put &inds.;
		%put --------------------------------------------------;
		%put;

		/*ファイルサイズの取得（（圧縮あり））*/
		ods output	EngineHost = _in_comp_enginehost(keep=Label1 nValue1 rename=(Label1=Label1_comp nValue1=nValue1_comp))
					Attributes = _in_comp_attributes(keep=Label2 nValue2);
		proc contents data = lib.&inds.
					  out  = _in_comp_cont_&i.
					  ;
		run;
		ods output close;

		/*現在のデータが圧縮済みか否かチェックする*/
		data _null_;
			set _in_comp_cont_&i.(obs=1);
			call symputx("compress",	compress);
			call symputx("in_comp_obs",	nobs);
			call symputx("modate",		modate);
		run;

		/*圧縮がかかっていない場合、圧縮して読み込む*/
		/*圧縮済みであれば、データを読み込む必要がなく、処理が速くなる*/
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

		/*変数の並び順を入力データと同じにする*/
		proc sort data = _in_comp_cont_&i.;
			by varnum;
		run;
		data _in_comp_cont_&i.;
			set _in_comp_cont_&i.;
			attrib
				format_length	length=$100. label="出力形式"
				informat_length	length=$100. label="入力形式"
			;
			if missing(format) then		format_length = "";
			else						format_length = cats(format, formatl, ".");
			if missing(informat) then	informat_length = "";
			else						informat_length = cats(informat, informl, ".");
		run;

		/*データの読込み（圧縮なし）*/
		/*
			圧縮なしでデータを読み込んだ時点で、ディスク不足エラーになることを回避するため、
			1000オブザベーション超は、ファイルサイズの推定値を保守的に算出し
			1000オブザベーション以下は実際のファイルサイズを取得する。

			ファイルサイズ（推定値） = オブザベーション数 × オブザベーションのバッファ長 × 1.02
			経験的に、実際のファイルサイズが1.02を乗じた場合の推定値を超える可能性は極めて低い。
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

		/*ファイルサイズの取得（（圧縮なし））*/
		ods output	EngineHost = _in_nocomp_enginehost(keep=Label1 nValue1);
		proc contents data = _in_nocomp;
		run;
		ods output close;

		/*ファイルサイズ一覧の作成*/
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

			if strip(Label2) in ("オブザベーション数", "変数の数", "オブザベーションのバッファ長") then do;
				id + 1;
				label = strip(Label2);
				value = nValue2;
				output;
			end;

			if strip(Label1) in ("ファイルサイズ (バイト)") then do;
				id + 1;
				label = "ファイルサイズGB（圧縮なし・推定値）";
				value = nValue1 / 1000000000;
				output;
			end;

			if strip(Label1_comp) in ("ファイルサイズ (バイト)") then do;
				id + 1;
				label = "ファイルサイズGB（圧縮あり）";
				value = nValue1_comp / 1000000000;
				output;
			end; 

		run;

		/*転置*/
		proc transpose data=_datasize_list_temp out=_datasize_list_tran(drop=_name_) prefix=var_;
			id id;
			idlabel label;
			var value;
		run;

		/*リスト*/
		data _datasize_list_&i.;
			attrib
				dir		length=$100.	label="ライブラリ名"
				dsname	length=$32.		label="データセット名"
				modate	length=8.		label="更新日"	format=datetime20.
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

			/*1000オブザベーション超の場合は推定値を算出する*/
			%if &in_comp_obs. > 1000 %then %do;
				var_4 = var_1 * var_3 * 1.02 / 1000000000;
			%end;

			format var_4 var_5 best.;
		run;

		/*不要データの削除*/
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

	/*集約*/
	data _datasize_list_fin;
		set _datasize_list_1 - _datasize_list_&dsnum.;
	run;


	/*エクセル出力*/
	ods excel file="%superq(excel)" 	options(
											sheet_name		="ファイルサイズ一覧"	/*シート名*/
											sheet_interval	="proc"					/*PROCごとにシートを分ける*/
											start_at		="2,2"					/*開始行列*/
											autofilter		="all"					/*オートフィルタ*/
											embedded_titles	="off"					/*タイトルなし*/
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

	/*不要データの削除*/
	proc datasets lib=work noprint;
		delete
			_datasize_list_:
		;
	quit;

	libname lib clear;

	options fmterr;

%mend MK_FILESIZE_LIST;


/*展開例1:フォルダ直下の全SASデータセット一覧の取得*/
/*
%MK_FILESIZE_LIST(
	dir		= C:\library\FCS,
	excel	= C:\Users\ファイルサイズ一覧\ファイルサイズ一覧_FCS.xlsx
);
*/

/*展開例2:フォルダ直下の指定したSASデータセット一覧の取得*/
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
	excel	=  C:\Users\ファイルサイズ一覧\ファイルサイズ一覧_DM.xlsx
);
*/

/*展開例3:ライブラリ参照名直下のSASデータセット一覧の取得*/
/*
%MK_FILESIZE_LIST(
	lib		= togo,
	excel	= C:\Users\ファイルサイズ一覧\ファイルサイズ一覧_統合マスタ.xlsx
);
*/
