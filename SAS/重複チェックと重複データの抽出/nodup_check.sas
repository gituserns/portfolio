%macro nodup_check(lib=, in=, key=_all_,  path=, name=);
%put --------------------------------------------------;
%put nodup_check;
%put 重複チェックと重複データの抽出;
%put &=lib;		/*ライブラリ参照先の指定（省略可）*/
%put &=in;		/*チェック対象データセットを指定（データセットオプションの指定可）*/
%put &=key;		/*チェック対象KEYを指定（既定値=_all_:全変数）*/
%put &=path;	/*ファイルの出力先を指定（末尾のバックスラッシュ"\"は不要）*/
%put &=name;	/*ファイル名を指定(ファイル名.html、ファイル名.xlsxが作成される)*/
%put --------------------------------------------------;

	/*フォーマットエラーを防ぐ*/
	options nofmterr compress=yes;

	/*ローカルマクロ変数の定義*/
	%local i num_indata num_dupdata;

	/*ライブラリ参照名の設定とデータの読み込み*/
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

	/*入力データのレコード数の取得*/
	%let num_indata = &sysnobs.;
	%if &num_indata. = 0 %then %put WARNING:入力データセットは空です。;

	/*変数名の取得*/
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

	/*変数のローカルマクロ変数化*/
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

	/*重複データの出力*/
	proc sort data=_nodup_check_data(keep=&key.) 
		out=_nodup_check_nodup 
		dupout=_nodup_check_dup
		nodupkey;
		by &key.;
	run;
	
	/*重複データの重複削除*/
	proc sort data=_nodup_check_dup nodupkey;
		by &key.;
	run;

	/*_nodup_check_dupのOBS数により条件分岐*/
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
				_dup_no 		length=8. 		label="重複No."
				_dup_key		length=$10000.	label="重複キー"
				_sabun_var		length=$10000.	label="差分のある変数"

				/*RETAIN用変数*/
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

			/*チェックキーの取得*/
			_dup_key = strip("&key.");

			/*重複レコードのログへの出力*/
			_logout_cnt + 1;
			if _logout_cnt <= 10 then put "WARNING:" _n_ +(-1) "オブザベーション目が重複しています。" +1 (&key.)(=);
			if _logout_cnt = 10 then put "WARNING:重複数が10に達したので、ログへの出力を止めます。";
			
			
			/*重複No.の取得*/
			if first.&&key&num_key.. then _dup_no + 1;

			/*差分のある変数の取得*/
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

			/*出力*/
			output _nodup_check_data_all;
			if last.&&key&num_key.. and not missing(_sabun_var) then output _nodup_check_sabun;

		run;

		/*差分のある変数の取得*/
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

		/*重複データのレコード数の取得*/
		%let num_dupdata = &sysnobs.;

		/*分布の集計*/
		proc freq data = _nodup_check_data_all_fin noprint;
			table _sabun_var / nocol norow nopercent out = _nodup_check_sabun_var_freq;
		run;

		/*整理*/
		data _nodup_check_sabun_var_freq_fin;
			attrib _dup_pattern	length=$200. label="重複パターン";
			set 
				_nodup_check_sabun_var_freq(in=in0 obs=1)
				_nodup_check_sabun_var_freq
				end = eof
			;
			if in0 then do;
				_dup_pattern = "重複なし";
				_sabun_var = "";
				count = sum(&num_indata., -&num_dupdata.);
				percent = count / &num_indata. * 100;
			end;
			else do;
				_dup_pattern = cats("重複パターン", _n_ - 1);
				percent = count / &num_indata. * 100;
			end;

			output;

			/*合計ステートメント*/
			_sum_count + count;
			_sum_percent + percent;
			
			if eof then do;
				_dup_pattern = "合計";
				_sabun_var = "";
				if _sum_count ne &num_indata. then put "WARNING:合計度数が入力データのレコード数と不一致"(_sum_count)(=);
				count = _sum_count;
				if not(99.99 < _sum_percent < 100.01) then put "WARNING:合計が100パーセントになっていない" (_sum_percent)(=);
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
		%put NOTE:重複データは存在しません。;
		data _nodup_check_data_all_fin;
			attrib
				_dup_key	length=$10000.	label="重複キー"
				_sabun_var	length=$20.	label="コメント"
			;

			/*チェックキーの取得*/
			_dup_key = strip("&key.");

			/*コメント*/
			_sabun_var = "重複レコードなし";
		run;

		data _nodup_check_sabun_var_freq_fin;
			attrib 
				_dup_pattern	length=$200.	label="重複パターン"
				count			length=8.		label="度数"
				percent			length=8.		label="合計度数のパーセント"

			;
			_dup_pattern = "重複なし";
			count = &num_indata.;
			percent = 100;
		run;

	%end;

	/*重複調査対象*/
	data _nodup_check_list;
		attrib
			komoku	length=$100.	label="項目"
			naiyo	length=$1000.	label="内容"
		;
		komoku	= "ライブラリ参照先";
		naiyo	= "%superq(lib)";
		output;

		komoku	= "テーブル名";
		naiyo	= "%superq(in)";
		output;

		komoku	= "調査キー";
		naiyo	= "%superq(key)";
		output;
	run;

	/*HTML出力*/
	ods html path = "%superq(path)" file ="%superq(name).html";
	title "重複調査";
	proc print data = _nodup_check_list label noobs;
	run;
	title "重複パターン分布";
	proc print data = _nodup_check_sabun_var_freq_fin label noobs;
	run;
	ods html close;
	title "";
	/*エクセル出力*/
	ods excel file="%superq(path)\%superq(name).xlsx" 	options(
									sheet_name		="重複パターン分布"	/*シート名*/
									sheet_interval	="none"				/*PROCごとにシートを分ける*/
									start_at		="2,2"				/*開始行列*/
									autofilter		="all"				/*オートフィルタ*/
									embedded_titles	="off"				/*タイトルなし*/
									);
	proc print data = _nodup_check_list label noobs;
	run;
	proc print data = _nodup_check_sabun_var_freq_fin label noobs;
	run;
/*
	ods excel options(
		sheet_name		="重複データ"
		sheet_interval	="proc"	
		start_at		="1,1"	
		autofilter		="all"	
		embedded_titles	="off"
	);
	proc print data=_nodup_check_data_all_fin label noobs;
	run;
*/
	ods excel close;

	/*重複データが多すぎるとメモリ不足を起こすため、EXPORTプロシジャで出力する*/
	proc export data=_nodup_check_data_all_fin
				outfile="%superq(path)\%superq(name).xlsx"
				dbms=xlsx
				label;
		sheet="重複データ"n;
	run;

	/*不要データの削除*/
	proc datasets lib=work noprint;
		delete _nodup_check_:;
	quit;

	options fmterr compress=no;

%mend nodup_check;


/*使用例*/
/*
%nodup_check(
	lib 	="C:\library",
	in		=D_0023_01(keep=TEN_CD OWNER_CD JIGYOSHA_NO TAX_KBN),
	key		=TEN_CD,
	path	=C:\Users\重複データ,
	name	=重複調査_D_0023_01
);

%nodup_check(
	lib		=,
	in		= D_0030_01_INVOICE,
	key		= TEN_CD KEIYAK_SDATE KEIYAK_EDATE,
	path	= C:\Users\重複データ,
	name	= 重複調査_D_0030_01_INVOICE
);

%nodup_check(
	lib		=,
	in		= D_0030_01_INVOICE_NODUP,
	key		= TEN_CD KEIYAK_SDATE KEIYAK_EDATE,
	path	= C:\Users\重複データ,
	name	= 重複調査_D_0030_01_INVOICE_NODUP
);
*/
