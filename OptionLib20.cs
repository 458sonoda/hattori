//
// OptionLibライブラリCustomクラス 
//
// このクラスについて
//	eFormLibライブラリのEWebFormクラスを利用して生成フォームへの
// アクセスを実現するためのクラススケルトンです。
// 用意されているpublic関数の機能をカスタマイズして利用します。
// 個々の関数についてはそれぞれの目的に応じて使用してください。
// 
// このクラスの生成では上位で生成されたEWebFormクラスとフォーム名を
// コンストラクタで与えてください。 
//		例：Custom custm = new Custom(eform,formname);
//
//
// -------------------------------------------------------------------------------------
// 下記のサンプルはフォームコントロールデータおよびプロパティへ
// アクセスするための例です。これらの機能等を利用してフォームへの
// アクセスを実現します。
//
// 下の例は"顧客名"というIDをもつテキスト項目のプロパティやデータへの
// アクセス方法です。このクラスの共通関数で利用します。
//	
// //「指定IDのコントロールデータオブジェクトへのアクセス例」
//		object obj = null;
//		bool bl = m_eForm.GetControleObject(ref obj,"顧客名");
//		string tp = obj.GetType().ToString();
//		System.Web.UI.WebControls.TextBox txt =	(System.Web.UI.WebControls.TextBox)obj;
//
//		// １．データ書き換え例
//		txt.Text = "abc";	
//
//		// ２．プロパティの設定変更例
//		txt.BackColor =  System.Drawing.Color.Red;
//
//		// ３．データの取得例
//		string str = txt.Text;	
//
//
//	// コントロールデータの書き換えは下記の関数も利用できます。
//  //************************************************************************
//  //画面項目の書き換え(項目名,値)
//  // ※1但し、mode=2[計算]の時は使用できません。mode=2[計算]の時は下記の※3の方法を使用して下さい。
//  //************************************************************************
//	 m_eForm.ChangeControlData("顧客名","abcde");
//
//  //************************************************************************
//  //読み取り専用項目の時はxmlの項目にも書き換えが必要です(xmlﾌｧｲﾙﾊﾟｽ,項目名,値)
//  //(""はｾｯﾄされない)
//  // ※2但し、mode=2[計算]の時は使用できません。mode=2[計算]の時は下記の※3の方法を使用して下さい。
//  //************************************************************************
//   m_eForm.ChangeXmlData_DOM(m_eForm.m_DataFile, "顧客名","abcde");     
//
//  //************************************************************************
//  //読み取り専用項目を一度に複数セットする場合は下記の要領で書き換えができます
//  // ※3mode=2[計算]の時は第4引数に1をセットします
//  //************************************************************************
//    EWebForm xform = new EWebForm();
//    xform.m_FileEncoding = m_eForm.m_FileEncoding;  //
//    xform.m_ReadXmlDataMode = 2;				      //Xmlデータ読込方式(2:DOM使用)
//    xform.XmlDataArray_Load(m_eForm.m_DataFile);	  //Xmlデータ読込(xmlﾌｧｲﾙﾊﾟｽ)
//    xform.XmlDataArray_Change("顧客名", "abcde", 1);//(項目名,値,空ﾃﾞｰﾀ書込ﾓｰﾄﾞ[1:書き込む 1以外:処理なし])
//    xform.XmlDataArray_Change("顧客名", "abcde", 1, 1);//(項目名,値,空ﾃﾞｰﾀ書込ﾓｰﾄﾞ[1:書き込む 1以外:処理なし], 1)※3
//    xform.XmlDataArray_Save();		              //Xmlデータ保存
//
//  // サンプル用共通関数
//  //************************************************************************
//  //画面データ取得
//  //************************************************************************
//    string ItemData = f_ItemDataGet("顧客名");
//
//  //************************************************************************
//  //ページ数取得
//  //************************************************************************
//    int PageCount = f_PageCountGet();
//
//  //************************************************************************
//  //シート数取得
//  // ※複数シートのときに指定ページ番号のシート数を返します
//  //************************************************************************
//    int SheetCount = f_SheetCountGet(PageNo);
//
//  //************************************************************************
//  //複数シート用項目名取得
//  // ※複数シートのときに指定ページ番号と指定シート番号に対応した項目名を返します
//  //************************************************************************
//    string ItemName = f_ItemNameGet(PageNo, SheetNo, "顧客名");
//
//  //------------------------------------------------------------------------
//  //------------------------------------------------------------------------
//  //サンプル説明
//  //------------------------------------------------------------------------
//  //------------------------------------------------------------------------
//    １．更新前処理(BeforeUpdate)で[売上伝票]フォームの明細データを更新
//        ・共通関数(f_ItemDataGet)より画面の[数量n]と[単価n]の値を取得しています。
//          m_eForm.GetControleObject(ref obj, ItemName);     //画面よりｺﾝﾄﾛｰﾙを取得
//          objのコントロールタイプにより値を取得
//        ・画面項目[金額n]に値をセットしています。
//          m_eForm.ChangeControlData("金額" + ino, i_kingaku);//書き込む項目名と値をｾｯﾄ
//        ・[金額n]が読み取り専用項目の場合は、データファイル(xml)に更新します。
//          下記のように書き込むデータファイルを読み込みます。
//            EWebForm xform = new EWebForm();
//            xform.m_FileEncoding = m_eForm.m_FileEncoding;  //上位PGで設定されているｴﾝｺｰﾄﾞをｾｯﾄ
//            xform.m_ReadXmlDataMode = 2;				      //Xmlデータ読込方式(2:DOM使用)
//            xform.XmlDataArray_Load(m_eForm.m_DataFile);	  //上位PGで設定されているxmlﾌｧｲﾙﾊﾟｽをｾｯﾄ
//          書き換える項目に値をセットします。
//            xform.XmlDataArray_Change("金額" + ino, i_kingaku, 1);//書き込む項目名と値をｾｯﾄ
//          最後にデータファイルに保存します。
//           xform.XmlDataArray_Save();		                  //Xmlデータ保存
//
//    ２．アクション処理(計算:mode=2)で[売上伝票]フォームの[日付]のチェック
//        ・共通関数(f_ItemDataGet)より画面の[日付]の値を取得しています。
//        ・取得した値が不正データの時、メッセージをセットしてリターンコードを返しています。
//          retstr = "この日付は入力できません。";
//          return -1;
//        ・更新前処理(BeforeUpdate)と同様に[売上伝票]フォームの明細データを更新しています。
//        　計算処理では直接画面の変更ができないため、画面項目[金額n]の値は変更していません。
//          書き換える項目に値をセットするときは、第4パラメータに1をセットします。
//            xform.XmlDataArray_Change("金額" + ino, i_kingaku, 1, 1);//書き込む項目名と値をｾｯﾄ
//
//    ３．アクション処理(入力画面表示時(参照作成ボタンから):mode=6)で[売上伝票]フォームの[伝票番号]のクリア
//        ・画面項目[伝票番号]に値をセットしています。
//          m_eForm.ChangeControlData("伝票番号","");         //書き込む項目名と値をｾｯﾄ
//        ・[伝票番号]が読み取り専用項目の場合は、データファイル(xml)に更新します。
//          m_eForm.ChangeXmlData_DOM(m_eForm.m_DataFile, "伝票番号", " ");     //(xmlﾌｧｲﾙﾊﾟｽ,項目名,値) ※この関数では""はｾｯﾄされないので半角ｽﾍﾟｰｽでｸﾘｱ
//
// -------------------------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;

using System.Diagnostics;
using System.IO;
using System.Collections.Specialized;

using System.Runtime.InteropServices;

using eFormLib20;
using System.Configuration;



//↑↑↑↑　ここまでは必須

//↓↓↓↓　ここから下はカスタマイズ内容に応じて追加
using System.Data.SqlClient;

using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualBasic;

namespace OptionLib20
{
    /// <summary>
    /// Custom のグローバル変数です。
    /// </summary>
    public static class GlobalVar
    {
        public static string strtemp = "";                                // debug
        public static double value = 0;                                   // debug

        public static Hashtable FORM = new Hashtable();     			  // FORM情報
        public static DataTable Csvtable = new DataTable("Table");

        // 新規フラグ//工程内検査表用に追加
        public static Dictionary<string, bool> NewDataFlg = new Dictionary<string, bool>();        
        // 申請フラグ
        public static Dictionary<string, string> ShinSeiFlg = new Dictionary<string, string>();
        // 却下フラグ
        public static Dictionary<string, bool> KyaKaFlg = new Dictionary<string, bool>();

    }

    /// <summary>
    /// Custom の概要の説明です。
    /// </summary>
    public class Custom
    {
        //------------------------------------------------------------------------------
        // クラス　グローバル変数
        //------------------------------------------------------------------------------
        // カスタム変数
        const string m_UserName = "TOSCO";
        const string m_VersionCode = "0.00";

        // 共通変数
        EWebForm m_eForm;           // eFormLibライブラリクラス
        Page m_Page;                // Page
        Panel m_Panel;              // Panel
        string m_FormName;          // Form XML File Name
        public int m_EffectFlag;    // 0:無効／1:有効
        string m_FormCD;            // FormCD
		string m_EditType;          // EDITTYPE
		int m_formID;               // FromID
        int m_Mode;                 // モード格納用
        int m_gensiQREnd = 0;       // 原紙QRコードの最終位置
        int m_filmQREnd = 0;        // フィルムQRコードの最終位置
        int m_flapQREnd = 0;        // フラップQRコードの最終位置

        string m_UserCode;          //ログインID//工程内検査表用に追加

        const string HEADCHAR = "_P+n1S+n2_";  // 複数ページ用ヘッダ文字定義 (+n1と+n2を置き換える)



        //画面ID
        const string 作液指示書_配合記録 = "HT001";
        const string 工程内検査表 = "HT002";
        const string 自己点検表 = "HT003";
        const string 品管用_作液指示書 = "HTMST";
        ///control.configのパス(服部製紙)
        //const string m_ControlConfig = "C:\\FormPat\\control.config"; //納品時はこちらを使用
        const string m_ControlConfig = "C:\\FormPat1\\control.config"; //納品時は削除
        const string m_hattoriConfig = "C:\\FormPat1\\hattori"; //納品時は削除
        const string m_DatabaseName = "database";

        //ログの出力最低レベル
        //　基本は1:INFO以上で出力。開発デバッグ時には最低レベルを0にする
        const int m_MinLogLevel = 0;

        //------------------------------------------------------------------------------
        // コンストラクタ
        //------------------------------------------------------------------------------
        // コンストラクタ(標準用１)
        public Custom(EWebForm ewebform, string fname, int eflag)
        {
            // コンストラクタ共通処理
            CustomCommon(ewebform, fname, eflag);
			SetSession();
		}
        // コンストラクタ(標準用２)
        public Custom(EWebForm ewebform, string fname)
        {	// コンストラクタ(標準用)
            // コンストラクタ共通処理
            CustomCommon(ewebform, fname, 1);
			SetSession();
		}
        // コンストラクタ（拡張用）
        public Custom()
        {
            m_EffectFlag = 1;		// クラス有効
			SetSession();
		}

        //------------------------------------------------------------------------------
        // コンストラクタ共通処理
        //------------------------------------------------------------------------------
        public void CustomCommon(EWebForm ewebform, string fname, int eflag)
        {
            m_EffectFlag = eflag;		// クラス有効
            m_eForm = ewebform;
            m_Page = ewebform.Pagem;
            m_Panel = ewebform.Panelm;
            m_FormName = fname;
            // 必要な初期処理はここに追加してください。
        }

		public void SetSession()
		{
            if (HttpContext.Current.Session["EDITTYPE"] == null)
			{
				m_EditType = "";
			}
			else
			{
                m_EditType = HttpContext.Current.Session["EDITTYPE"].ToString();
			}
            if (HttpContext.Current.Session["UserCode"] == null)
            {
                m_UserCode = "";

            }
            if (HttpContext.Current.Session["formcd"] == null)
            {
                m_FormCD = "";

            }
            else
            {
                m_UserCode = HttpContext.Current.Session["UserCode"].ToString();

               
            }
            if (HttpContext.Current.Session["FormID"] == null)
			{

				m_formID = 0;
			}
			else
			{
                m_formID = Convert.ToInt32(HttpContext.Current.Session["FormID"].ToString());
			}
		}

        //------------------------------------------------------------------------------
        // 基本関数
        //------------------------------------------------------------------------------
        // このライブラリのカスタマイズユーザー名取得
        public string GetUserName()
        {
            return m_UserName;
        }
        // このライブラリのカスタマイズユーザー設定バージョン取得
        public string GetVersionCode()
        {
            return m_VersionCode;
        }
        // フォームCDの設定
        public void SetFormCD(string fcode)
        {
            m_FormCD = fcode.Trim();
        }

        /// <summary>
        /// 2021年2月８日を文字列2021/02/08に変更。
        /// </summary>
        private string CalenderToStringDate(string nengetubi)
        {
            string retstr = string.Empty;

            string[] tenkenbiar = nengetubi.Split('/');
            if (nengetubi.Length == 10 || nengetubi.Length == 9 || nengetubi.Length == 8)
            {
                string y = tenkenbiar[0];
                string m = tenkenbiar[1];
                string d = tenkenbiar[2];
                // 月、日に前ゼロ付加
                retstr = y + "/" + m.PadLeft(2, '0') + "/" + d.PadLeft(2, '0');
            }
            //年月の取得
            else if (nengetubi.Length == 7 || nengetubi.Length == 6)
            {
                string y = tenkenbiar[0];
                string m = tenkenbiar[1];
                // 月、日に前ゼロ付加
                retstr = y + "/" + m.PadLeft(2, '0');
            }
            return retstr;
        }

        /// <summary>
        /// 工程内検査表入力チェック処理
        /// <param name="renban">連番</param>
        /// </summary>
        /// <returns>retstr</returns>
        private string checkItemKoteinaikensa(ref string renban)                                                 
        {
            renban = f_ItemDataGet("renban");
            string retstr = string.Empty;
           
            if(renban == string.Empty)
            {
                retstr = "連番が記載されていません"; //ここのエラーメッセージは変更すること
                return retstr;
            }
            
            return retstr;
        }

        /// <summary>
        /// 作液指示書配合記録_品管用入力チェック処理
        /// <param name="seihinmei">製品名</param>
        /// <param name="yakuekimei">薬液名</param>
        /// </summary>
        /// <returns>retstr</returns>
        private string checkItemSakuekiHinkan(ref string seihinmei,string yakuekimei)
        {
           
            string retstr = string.Empty;
            //double pH1 = 0;
            //double pH2 = 0;
            //double hijyu1 = 0;
            //double hijyu2 = 0;

            if (seihinmei == string.Empty && yakuekimei == string.Empty)
            {
                retstr = "製品名及び薬液名が記載されていません。"; 
                return retstr;
            }
            if (seihinmei == string.Empty && yakuekimei != string.Empty)
            {
                retstr = "製品名が記載されていません。";
                return retstr;
            }
            if (seihinmei != string.Empty && yakuekimei == string.Empty)
            {
                retstr = "薬液名が記載されていません。";
                return retstr;
            }
            //規格欄の大小が入れ替わってないか確認
            //if (f_ItemDataGet("規格4") != string.Empty && f_ItemDataGet("規格5") != string.Empty)
            //{
            //    pH1 = double.Parse( f_ItemDataGet("規格4"));
            //    pH2 = double.Parse(f_ItemDataGet("規格5"));
            //    if(pH1 > pH2)
            //    {
            //        retstr = "pHの範囲の最大値と最小値が入れ替わっています。";
            //        return retstr;
            //    }               
            //}
            //if (f_ItemDataGet("規格6") != string.Empty && f_ItemDataGet("規格7") != string.Empty)
            //{
            //    hijyu1 = double.Parse(f_ItemDataGet("規格6"));
            //    hijyu2 = double.Parse(f_ItemDataGet("規格6"));
            //    if (hijyu1 > hijyu2)
            //    {
            //        retstr = "比重の範囲の最大値と最小値が入れ替わっています。";
            //        return retstr;
            //    }
            //}
            return retstr;
        }


        /// <summary>
        /// config.contorolからデータベース名を取得する
        /// </summary>
        /// <returns>データベース名</returns>
        private string getDbName()
        {
            //control.configよりデータベースの接続先を取得する
            XmlDocument control = new XmlDocument();
            control.Load(m_ControlConfig);
            string strdatabase = control.GetElementsByTagName(m_DatabaseName)[0].InnerText;
            return strdatabase;
        }
        /// <summary>
        /// CSVデータをDBに登録
        /// </summary>
        /// <returns>項目定義の有無(true:定義あり false:定義なし</returns>
        private string GetQRCode(string itemnameS, string itemnameD, int start, int end)
        {
            string retstr = string.Empty;
            string itemname=string.Empty;
            string qrs;
            //xxxxx_eFormLibTemp.Xmlを読み込む
            EWebForm xform = new EWebForm();
            xform.m_FileEncoding = m_eForm.m_FileEncoding;  //
            xform.m_ReadXmlDataMode = 2;                    //Xmlデータ読込方式(2:DOM使用)
            xform.XmlDataArray_Load(m_eForm.m_DataFile);    //Xmlデータ読込

            // 現在セットされているＱＲコードを確定する(Action mode 2のみ)(手で消したとき対応)
            if (m_Mode == 2)
            {
                setQRCodeNow(itemnameD, end, xform);
            }

            // QR取得
            qrs = Strings.StrConv(f_ItemDataGet(itemnameS), VbStrConv.Narrow);
            if (qrs == string.Empty)
            {
                //retstr = "QRコードなし";
                return retstr;
            }
            
            //改行コードで分割
            string[] qr = qrs.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            //retstr = qr[0];

            for (int i = start, j=1; j <= qr.Length; i++, j++)
            {
                //カンマ区切りで分割    
                string[] qrcode = qr[j - 1].Split(',');

                retstr = qrcode[0];

                itemname = itemnameD + "_" + i;

                // File.WriteAllText(@"C:\FormPatData1\temp\QR_" + itemnameD  + "_" + j + ".txt", "(" + m_Mode + ")(" + i + ")" + retstr);
                if (m_Mode == 2)
                {
                    xform.XmlDataArray_Change(itemname, qrcode[0], 1, 1);       // Action時更新用
                }
                else
                {
                    m_eForm.ChangeControlData(itemname, qrcode[0]);             // BeforeUpadate時更新用
                }

            }
            xform.XmlDataArray_Save();      //xxxxx_eFormLibTemp.Xmlに保存

            return retstr;

        }

        
        /// <summary>
        /// QRコードの次の追加位置を取得
        /// </summary>
        /// <returns>項目定義の有無(true:定義あり false:定義なし</returns>
        private int getGensiQRStart(string target,int iEnd)
        {
            int iRet = iEnd;
            string QRcode = string.Empty;

            for (int i = iEnd; i >= 0; i--)
            {
                QRcode = f_ItemDataGet(target + "_" + i);
                if(!QRcode.Equals("") || (i==0))
                {
                    iRet = i+1;
                    break;
                }
            }

            return iRet;
        }

        /// <summary>
        /// CSVデータをDBに登録
        /// </summary>
        /// <returns>項目定義の有無(true:定義あり false:定義なし</returns>
        private void setQRCodeNow(string itemnameD, int end, EWebForm xform)
        {
            string QRcode = string.Empty;
            string itemname = string.Empty;

            for (int i=1; i<=end; i++)
            {
                itemname = itemnameD + "_" + i;
                QRcode = f_ItemDataGet(itemname);
                if (QRcode.Equals(""))
                {
                    //File.WriteAllText(@"C:\FormPatData1\temp\QR_" + itemname + ".txt", "(" + m_Mode + ")(" + itemname + ")[" + " " + "]");
                    xform.XmlDataArray_Change(itemname, " ", 1, 1);       // Action時更新用
                }
            }
            xform.XmlDataArray_Save();      //xxxxx_eFormLibTemp.Xmlに保存
        }


            /// <summary>
            /// 工程内検査表を登録する。
            /// <returns></returns>
            private string registKoteinaikensa(string renban)
        {
            string retstr = "";
            string function_name = "f_registKoteinaikensa";


            ////製造指示書No.を発番         
            //string hizyke = DateTime.Now.ToShortDateString();//yyyy/mm/dd
            //string hizeke_d8 = hizyke.Replace("/", "");//yyyymmdd

            //○年○月○日の日付をyyyyMMddに変換
            string hiduke = f_ItemDataGet("日付");
            if (hiduke != "") {
                hiduke = f_ItemDataGet("日付").Replace("年", "/").Replace("月", "/").Replace("日", "");
                DateTime dt = DateTime.Parse(hiduke);
                hiduke = dt.ToString("yyyyMMdd");
            }
                       

            //DBの号機項目に画面上で選択された号機番号を登録する用の変数
            string gouki_1 = f_ItemDataGet("号機_1");
            string gouki_2 = f_ItemDataGet("号機_2");

            //工程内検査結果1～3
            string goukaku1 = f_ItemDataGet("合格ラジオボタン1");
            string fugokaku1 = f_ItemDataGet("不合格ラジオボタン1");
            string goukaku2 = f_ItemDataGet("合格ラジオボタン2");
            string fugokaku2 = f_ItemDataGet("不合格ラジオボタン2");
            string goukaku3 = f_ItemDataGet("合格ラジオボタン3");
            string fugokaku3 = f_ItemDataGet("不合格ラジオボタン3");

            string kensakekka1 = string.Empty;
            string kensakekka2 = string.Empty;
            string kensakekka3 = string.Empty;

            //号機
            string gouki_num = string.Empty;
            if (gouki_1.Equals("True"))
            {
                gouki_num = "1";
            }
            else if (gouki_2.Equals("True"))
            {
                gouki_num = "2";
            }
            

            //工程内検査結果           
            if (goukaku1.Equals("True"))
            {
                kensakekka1 = "合格";
            }
            else if (fugokaku1.Equals("True"))
            {
                kensakekka1 = "不合格";
            }

            if (goukaku2.Equals("True"))
            {
                kensakekka2 = "合格";
            }
            else if (fugokaku2.Equals("True"))
            {
                kensakekka2 = "不合格";
            }

            if (goukaku3.Equals("True"))
            {
                kensakekka3 = "合格";
            }
            else if (fugokaku3.Equals("True"))
            {
                kensakekka3 = "不合格";
            }
           
            //// データベース名の取得
            //string strdatabase = getDbName();

            //control.configよりデータベースの接続先を取得する
            XmlDocument control = new XmlDocument();
            control.Load(m_ControlConfig);
            string strdatabase = control.GetElementsByTagName(m_DatabaseName)[0].InnerText;
           

            using (var conn = new SqlConnection(strdatabase))
            {
                try
                {
                    // DB接続
                    var command = conn.CreateCommand();
                    conn.Open();

                    //トランザクションの開始
                    SqlTransaction transaction = conn.BeginTransaction();

                    command.Transaction = transaction;

                    // 削除
                    //command.CommandText = "DELETE FROM 服部製紙様用_工程内検査表 WHERE 連番 = '" + renban + "'";

#if DEBUG




#endif                    
                    try
                    {
                        command.CommandText = "INSERT INTO 服部製紙様用_工程内検査表 ( " +
                            "[連番]," +
                            "[号機]," +
                            "[日付]," +
                            "[製品名]," +
                            "[記録責任者1]," +
                            "[記録責任者2]," +
                            "[記録責任者3]," +
                            "[生産予定数量1]," +
                            "[製品枚数]," +
                            "[生産時間1]," +
                            "[生産時間2]," +
                            "[生産時間3]," +
                            "[生産時間4]," +
                            "[生産時間5]," +
                            "[生産時間6]," +
                            "[製品ロットNo1]," +
                            "[製品ロットNo2]," +
                            "[製品ロットNo3]," +
                            "[薬液ロットNo1]," +
                            "[薬液ロットNo2]," +
                            "[薬液ロットNo3]," +
                            "[液量1]," +
                            "[液量2]," +
                            "[液量3]," +
                            "[オペレータ1]," +
                            "[オペレータ2]," +
                            "[オペレータ3]," +
                            "[検品者1]," +
                            "[検品者2]," +
                            "[検品者3]," +
                            "[梱包者1]," +
                            "[梱包者2]," +
                            "[梱包者3]," +
                            "[原紙名]," +
                            "[原紙ロットNo1_1]," +
                            "[原紙ロットNo1_2]," +
                            "[原紙ロットNo1_3]," +
                            "[原紙ロットNo1_4]," +
                            "[原紙ロットNo1_5]," +
                            "[原紙ロットNo1_6]," +
                            "[原紙ロットNo1_7]," +
                            "[原紙ロットNo1_8]," +
                            "[原紙ロットNo1_9]," +
                            "[原紙ロットNo1_10]," +
                            "[原紙ロットNo1_11]," +
                            "[原紙ロットNo1_12]," +
                            "[原紙ロットNo1_13]," +
                            "[原紙ロットNo1_14]," +
                            "[原紙ロットNo1_15]," +
                            "[原紙ロットNo1_16]," +
                            "[原紙ロットNo1_17]," +
                            "[原紙ロットNo1_18]," +
                            "[原紙ロットNo1_19]," +
                            "[原紙ロットNo1_20]," +
                            "[原紙ロットNo1_21]," +
                            "[原紙ロットNo1_22]," +
                            "[原紙ロットNo1_23]," +
                            "[原紙ロットNo1_24]," +
                            "[原紙ロットNo1_25]," +
                            "[原紙ロットNo1_26]," +
                            "[原紙ロットNo1_27]," +
                            "[原紙ロットNo1_28]," +
                            "[フィルムロットNo1_1]," +
                            "[フィルムロットNo1_2]," +
                            "[フィルムロットNo1_3]," +
                            "[フィルムロットNo2_1]," +
                            "[フィルムロットNo2_2]," +
                            "[フィルムロットNo2_3]," +
                            "[フィルムロットNo3_1]," +
                            "[フィルムロットNo3_2]," +
                            "[フィルムロットNo3_3]," +
                            "[フィルムロットNo4_1]," +
                            "[フィルムロットNo4_2]," +
                            "[フィルムロットNo4_3]," +
                            "[フィルムロットNo_継目1_1]," +
                            "[フィルムロットNo_継目1_2]," +
                            "[フィルムロットNo_継目1_3]," +
                            "[フィルムロットNo_継目2_1]," +
                            "[フィルムロットNo_継目2_2]," +
                            "[フィルムロットNo_継目2_3]," +
                            "[フィルムロットNo_継目3_1]," +
                            "[フィルムロットNo_継目3_2]," +
                            "[フィルムロットNo_継目3_3]," +
                            "[フィルムロットNo_継目4_1]," +
                            "[フィルムロットNo_継目4_2]," +
                            "[フィルムロットNo_継目4_3]," +
                            "[フラップロットNo1]," +
                            "[フラップロットNo2]," +
                            "[フラップロットNo3]," +
                            "[印刷状態1]," +
                            "[印刷状態2]," +
                            "[印刷状態3]," +
                            "[位置差異1]," +
                            "[位置差異2]," +
                            "[位置差異3]," +
                            "[ストッパー抜き1]," +
                            "[ストッパー抜き2]," +
                            "[ストッパー抜き3]," +
                            "[ロット印字内容1]," +
                            "[ロット印字内容2]," +
                            "[ロット印字内容3]," +
                            "[入数1]," +
                            "[入数2]," +
                            "[入数3]," +
                            "[原紙カウンター数1]," +
                            "[原紙カウンター数2]," +
                            "[原紙カウンター数3]," +
                            "[包装機カウンター数1]," +
                            "[包装機カウンター数2]," +
                            "[包装機カウンター数3]," +
                            "[ラベラーカウンター数1]," +
                            "[ラベラーカウンター数2]," +
                            "[ラベラーカウンター数3]," +
                            "[実生産個数1]," +
                            "[実生産個数2]," +
                            "[実生産個数3]," +
                            "[内ケース使用数1]," +
                            "[ロス1]," +
                            "[内ケース使用数2]," +
                            "[ロス3]," +
                            "[内ケース使用数3]," +
                            "[ロス5]," +
                            "[外ケース使用数1]," +
                            "[ロス2]," +
                            "[外ケース使用数2]," +
                            "[ロス4]," +
                            "[外ケース使用数3]," +
                            "[ロス6]," +
                            "[原紙歩留1]," +
                            "[原紙歩留2]," +
                            "[原紙歩留3]," +
                            "[フィルム歩留1]," +
                            "[フィルム歩留2]," +
                            "[フィルム歩留3]," +
                            "[原紙ロス1]," +
                            "[原紙ロス2]," +
                            "[原紙ロス3]," +
                            "[フラップロス1]," +
                            "[フラップロス2]," +
                            "[フラップロス3]," +
                            "[フィルムロス1]," +
                            "[フィルムロス2]," +
                            "[フィルムロス3]," +
                            "[フラップ異常1]," +
                            "[フラップ異常2]," +
                            "[フラップ異常3]," +
                            "[原紙噛み込み1]," +
                            "[原紙噛み込み2]," +
                            "[原紙噛み込み3]," +
                            "[フィルム交換1]," +
                            "[フィルム交換2]," +
                            "[フィルム交換3]," +
                            "[パックレス1]," +
                            "[パックレス2]," +
                            "[パックレス3]," +
                            "[検査ロス1]," +
                            "[検査ロス2]," +
                            "[検査ロス3]," +
                            "[調整ロス1]," +
                            "[調整ロス2]," +
                            "[調整ロス3]," +
                            "[予備1]," +
                            "[予備2]," +
                            "[予備3]," +
                            "[工程内検査結果1]," +
                            "[工程内検査結果2]," +
                            "[工程内検査結果3]," +
                            "[備考1]," +
                            "[備考2]," +
                            "[備考3]," +
                            "[判定者名]" +
                            ") VALUES ('" +
                            //f_ItemDataGet("renban") + "','" +
                            renban + "','" +
                            gouki_num + "','" +
                            hiduke + "','" +
                            f_ItemDataGet("製品名") + "','" +
                            f_ItemDataGet("記録責任者1") + "','" +
                            f_ItemDataGet("記録責任者2") + "','" +
                            f_ItemDataGet("記録責任者3") + "','" +
                            f_ItemDataGet("生産予定数量1") + "','" +
                            f_ItemDataGet("製品枚数") + "','" +
                            f_ItemDataGet("生産時間1") + "','" +
                            f_ItemDataGet("生産時間2") + "','" +
                            f_ItemDataGet("生産時間3") + "','" +
                            f_ItemDataGet("生産時間4") + "','" +
                            f_ItemDataGet("生産時間5") + "','" +
                            f_ItemDataGet("生産時間6") + "','" +
                            f_ItemDataGet("製品ロットNo1") + "','" +
                            f_ItemDataGet("製品ロットNo2") + "','" +
                            f_ItemDataGet("製品ロットNo3") + "','" +
                            f_ItemDataGet("薬液ロットNo1") + "','" +
                            f_ItemDataGet("薬液ロットNo2") + "','" +
                            f_ItemDataGet("薬液ロットNo3") + "','" +
                            f_ItemDataGet("液量1") + "','" +
                            f_ItemDataGet("液量2") + "','" +
                            f_ItemDataGet("液量3") + "','" +
                            f_ItemDataGet("オペレータ1") + "','" +
                            f_ItemDataGet("オペレータ2") + "','" +
                            f_ItemDataGet("オペレータ3") + "','" +
                            f_ItemDataGet("検品者1") + "','" +
                            f_ItemDataGet("検品者2") + "','" +
                            f_ItemDataGet("検品者3") + "','" +
                            f_ItemDataGet("梱包者1") + "','" +
                            f_ItemDataGet("梱包者2") + "','" +
                            f_ItemDataGet("梱包者3") + "','" +
                            f_ItemDataGet("原紙名") + "','" +
                            f_ItemDataGet("原紙ロットNo1_1") + "','" +
                            f_ItemDataGet("原紙ロットNo1_2") + "','" +
                            f_ItemDataGet("原紙ロットNo1_3") + "','" +
                            f_ItemDataGet("原紙ロットNo1_4") + "','" +
                            f_ItemDataGet("原紙ロットNo1_5") + "','" +
                            f_ItemDataGet("原紙ロットNo1_6") + "','" +
                            f_ItemDataGet("原紙ロットNo1_7") + "','" +
                            f_ItemDataGet("原紙ロットNo1_8") + "','" +
                            f_ItemDataGet("原紙ロットNo1_9") + "','" +
                            f_ItemDataGet("原紙ロットNo1_10") + "','" +
                            f_ItemDataGet("原紙ロットNo1_11") + "','" +
                            f_ItemDataGet("原紙ロットNo1_12") + "','" +
                            f_ItemDataGet("原紙ロットNo1_13") + "','" +
                            f_ItemDataGet("原紙ロットNo1_14") + "','" +
                            f_ItemDataGet("原紙ロットNo1_15") + "','" +
                            f_ItemDataGet("原紙ロットNo1_16") + "','" +
                            f_ItemDataGet("原紙ロットNo1_17") + "','" +
                            f_ItemDataGet("原紙ロットNo1_18") + "','" +
                            f_ItemDataGet("原紙ロットNo1_19") + "','" +
                            f_ItemDataGet("原紙ロットNo1_20") + "','" +
                            f_ItemDataGet("原紙ロットNo1_21") + "','" +
                            f_ItemDataGet("原紙ロットNo1_22") + "','" +
                            f_ItemDataGet("原紙ロットNo1_23") + "','" +
                            f_ItemDataGet("原紙ロットNo1_24") + "','" +
                            f_ItemDataGet("原紙ロットNo1_25") + "','" +
                            f_ItemDataGet("原紙ロットNo1_26") + "','" +
                            f_ItemDataGet("原紙ロットNo1_27") + "','" +
                            f_ItemDataGet("原紙ロットNo1_28") + "','" +
                            f_ItemDataGet("フィルムロットNo1_1") + "','" +
                            f_ItemDataGet("フィルムロットNo1_2") + "','" +
                            f_ItemDataGet("フィルムロットNo1_3") + "','" +
                            f_ItemDataGet("フィルムロットNo2_1") + "','" +
                            f_ItemDataGet("フィルムロットNo2_2") + "','" +
                            f_ItemDataGet("フィルムロットNo2_3") + "','" +
                            f_ItemDataGet("フィルムロットNo3_1") + "','" +
                            f_ItemDataGet("フィルムロットNo3_2") + "','" +
                            f_ItemDataGet("フィルムロットNo3_3") + "','" +
                            f_ItemDataGet("フィルムロットNo4_1") + "','" +
                            f_ItemDataGet("フィルムロットNo4_2") + "','" +
                            f_ItemDataGet("フィルムロットNo4_3") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目1_1") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目1_2") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目1_3") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目2_1") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目2_2") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目2_3") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目3_1") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目3_2") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目3_3") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目4_1") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目4_2") + "','" +
                            f_ItemDataGet("フィルムロットNo_継目4_3") + "','" +
                            f_ItemDataGet("フラップロットNo1") + "','" +
                            f_ItemDataGet("フラップロットNo2") + "','" +
                            f_ItemDataGet("フラップロットNo3") + "','" +
                            f_ItemDataGet("印刷状態1") + "','" +
                            f_ItemDataGet("印刷状態2") + "','" +
                            f_ItemDataGet("印刷状態3") + "','" +
                            f_ItemDataGet("位置差異1") + "','" +
                            f_ItemDataGet("位置差異2") + "','" +
                            f_ItemDataGet("位置差異3") + "','" +
                            f_ItemDataGet("ストッパー抜き1") + "','" +
                            f_ItemDataGet("ストッパー抜き2") + "','" +
                            f_ItemDataGet("ストッパー抜き3") + "','" +
                            f_ItemDataGet("ロット印字内容1") + "','" +
                            f_ItemDataGet("ロット印字内容2") + "','" +
                            f_ItemDataGet("ロット印字内容3") + "','" +
                            f_ItemDataGet("入数1") + "','" +
                            f_ItemDataGet("入数2") + "','" +
                            f_ItemDataGet("入数3") + "','" +
                            f_ItemDataGet("原紙カウンター数1") + "','" +
                            f_ItemDataGet("原紙カウンター数2") + "','" +
                            f_ItemDataGet("原紙カウンター数3") + "','" +
                            f_ItemDataGet("包装機カウンター数1") + "','" +
                            f_ItemDataGet("包装機カウンター数2") + "','" +
                            f_ItemDataGet("包装機カウンター数3") + "','" +
                            f_ItemDataGet("ラベラーカウンター数1") + "','" +
                            f_ItemDataGet("ラベラーカウンター数2") + "','" +
                            f_ItemDataGet("ラベラーカウンター数3") + "','" +
                            f_ItemDataGet("実生産個数1") + "','" +
                            f_ItemDataGet("実生産個数2") + "','" +
                            f_ItemDataGet("実生産個数3") + "','" +
                            f_ItemDataGet("内ケース使用数1") + "','" +
                            f_ItemDataGet("ロス1") + "','" +
                            f_ItemDataGet("内ケース使用数2") + "','" +
                            f_ItemDataGet("ロス3") + "','" +
                            f_ItemDataGet("内ケース使用数3") + "','" +
                            f_ItemDataGet("ロス5") + "','" +
                            f_ItemDataGet("外ケース使用数1") + "','" +
                            f_ItemDataGet("ロス2") + "','" +
                            f_ItemDataGet("外ケース使用数2") + "','" +
                            f_ItemDataGet("ロス4") + "','" +
                            f_ItemDataGet("外ケース使用数3") + "','" +
                            f_ItemDataGet("ロス6") + "','" +
                            f_ItemDataGet("原紙歩留1") + "','" +
                            f_ItemDataGet("原紙歩留2") + "','" +
                            f_ItemDataGet("原紙歩留3") + "','" +
                            f_ItemDataGet("フィルム歩留1") + "','" +
                            f_ItemDataGet("フィルム歩留2") + "','" +
                            f_ItemDataGet("フィルム歩留3") + "','" +
                            f_ItemDataGet("原紙ロス1") + "','" +
                            f_ItemDataGet("原紙ロス2") + "','" +
                            f_ItemDataGet("原紙ロス3") + "','" +
                            f_ItemDataGet("フラップロス1") + "','" +
                            f_ItemDataGet("フラップロス2") + "','" +
                            f_ItemDataGet("フラップロス3") + "','" +
                            f_ItemDataGet("フィルムロス1") + "','" +
                            f_ItemDataGet("フィルムロス2") + "','" +
                            f_ItemDataGet("フィルムロス3") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_1") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_1") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_1") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_2") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_2") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_2") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_3") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_3") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_3") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_4") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_4") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_4") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_5") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_5") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_5") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_6") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_6") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_6") + "','" +
                            f_ItemDataGet("不具合の内容と詳細_7") + "','" +
                            f_ItemDataGet("不具合の内容と詳細1_7") + "','" +
                            f_ItemDataGet("不具合の内容と詳細2_7") + "','" +
                            kensakekka1 + "','" +
                            kensakekka2 + "','" +
                            kensakekka3 + "','" +                           
                            f_ItemDataGet("備考1") + "','" +
                            f_ItemDataGet("備考2") + "','" +
                            f_ItemDataGet("備考3") + "','" +
                            f_ItemDataGet("判定者名")  + "'" +
                            ")";
                        
#if DEBUG
                        f_LogWrite(3, retstr + command.CommandText);
#endif
                        command.ExecuteNonQuery();

                    }
                    catch (SqlException ex)
                    {
                        switch (ex.Number)
                        {
                            case 2627:
                            case -2:
                                retstr = "二重登録エラーが発生しました。(" + function_name + ")";

                                break;
                            default:
                                retstr = "データベースエラーが発生しました。1(" + ex.Message + "[" + function_name + "])";
                                break;

                        }
                        f_LogWrite(3, retstr + "[" + ex.Message + "]");
                        //トランザクションのロールバック
                        transaction.Rollback();
                        return retstr;
                    }
                    catch (Exception ex)
                    {
                        retstr = "例外エラーが発生しました。1(" + ex.Message + "[" + function_name + "])";
                        return retstr;
                        //throw;
                    }

                    // トランザクションのコミット
                    transaction.Commit();

                    conn.Close();
                    conn.Dispose();
                }
                catch (Exception ex)
                {
                    retstr = "例外エラーが発生しました。2(" + function_name + ")";
                    f_LogWrite(3, retstr + "[" + ex.Message + "]" + "[" + ex.StackTrace + "]");
                    return retstr;
                }
                finally
                {

                }
            }

            return retstr;
        }

        /// <summary>
        /// 作液指示書_品管用を登録する。
        /// <returns></returns>
        private string registSakuekihaigoHinkan(string seihinmei,string yakuekimei)
        {
            string retstr = "";
            string function_name = "f_registSakuekihaigoHinkan";


            //// データベース名の取得
            //string strdatabase = getDbName();

            //control.configよりデータベースの接続先を取得する
            XmlDocument control = new XmlDocument();
            control.Load(m_ControlConfig);

#if Azure
            //下のはアジュールの為にやっている.。
            string strdatabase = "server=toscoformpat.database.windows.net;uid=toscowest;pwd=tosco@75;Initial Catalog=FormPat";
#else
            //下のはBD3823215082の為にやっている。
            string strdatabase = control.GetElementsByTagName(m_DatabaseName)[0].InnerText;
#endif

            using (var conn = new SqlConnection(strdatabase))
            {
                try
                {
                    // DB接続
                    var command = conn.CreateCommand();
                    conn.Open();

                    //トランザクションの開始
                    SqlTransaction transaction = conn.BeginTransaction();

                    command.Transaction = transaction;

                    // 重複するものは削除
                    command.CommandText = "DELETE FROM 服部製紙様用_作液指示書_品管用 " 
                                               + "WHERE ログインID = '" + m_UserCode + "'"
                                               + "  AND     製品名 = '" + seihinmei   + "'" 
                                               + "  AND     薬液名 = '" + yakuekimei + "'";

#if DEBUG

                    f_LogWrite(3, retstr + command.CommandText + "(" + strdatabase + ")");

#endif
                    command.ExecuteNonQuery();
                    try
                    {
                        command.CommandText = "INSERT INTO 服部製紙様用_作液指示書_品管用 ( " +

                            "[ログインID]," +
                            "[製品名]," +
                            "[薬液名]," +
                            "[処方番号]," +
                            "[薬品名1]," +
                            "[薬品名2]," +
                            "[薬品名3]," +
                            "[薬品名4]," +
                            "[薬品名5]," +
                            "[薬品名6]," +
                            "[薬品名7]," +
                            "[薬品名8]," +
                            "[薬品名9]," +
                            "[薬品名10]," +
                            "[薬品名11]," +
                            "[薬品名12]," +
                            "[薬品名13]," +
                            "[薬品名14]," +
                            "[薬品名15]," +
                            "[薬品名16]," +
                            "[薬品名17]," +
                            "[薬品名18]," +
                            "[薬品名19]," +
                            "[薬品名20]," +
                            "[薬品名21]," +
                            "[薬品名22]," +
                            "[薬品名23]," +
                            "[薬品名24]," +
                            "[薬品名25]," +
                            "[薬品名26]," +
                            "[薬品名27]," +
                            "[薬品名28]," +
                            "[薬品名29]," +
                            "[薬品名30]," +
                            "[作液工程名1]," +
                            "[作液工程名2]," +
                            "[作液工程名3]," +
                            "[作液工程名4]," +
                            "[作液工程名5]," +
                            "[作液工程名6]," +
                            "[作液工程名7]," +
                            "[作液工程名8]," +
                            "[作液工程名9]," +
                            "[作液工程名10]," +
                            "[作液工程名11]," +
                            "[作液工程名12]," +
                            "[作液工程名13]," +
                            "[作液工程名14]," +
                            "[作液工程名15]," +
                            "[作液工程名16]," +
                            "[作液工程名17]," +
                            "[作液工程名18]," +
                            "[作液工程名19]," +
                            "[作液工程名20]," +
                            "[作液工程名21]," +
                            "[作液工程名22]," +
                            "[作液工程名23]," +
                            "[作液工程名24]," +
                            "[作液工程名25]," +
                            "[作液工程名26]," +
                            "[作液工程名27]," +
                            "[作液工程名28]," +
                            "[作液工程名29]," +
                            "[作液工程名30]," +
                            "[配合率1]," +
                            "[配合率2]," +
                            "[配合率3]," +
                            "[配合率4]," +
                            "[配合率5]," +
                            "[配合率6]," +
                            "[配合率7]," +
                            "[配合率8]," +
                            "[配合率9]," +
                            "[配合率10]," +
                            "[配合率11]," +
                            "[配合率12]," +
                            "[配合率13]," +
                            "[配合率14]," +
                            "[配合率15]," +
                            "[配合率16]," +
                            "[配合率17]," +
                            "[配合率18]," +
                            "[配合率19]," +
                            "[配合率20]," +
                            "[配合率21]," +
                            "[配合率22]," +
                            "[配合率23]," +
                            "[配合率24]," +
                            "[配合率25]," +
                            "[配合率26]," +
                            "[配合率27]," +
                            "[配合率28]," +
                            "[配合率29]," +
                            "[配合率30]," +
                            "[備考1]," +
                            "[備考2]," +
                            "[規格1]," +
                            "[規格2]," +
                            "[規格3]," +
                            "[規格4]," +
                            "[規格5]," +
                            "[規格6]," +
                            "[規格7]," +
                            "[規格8]," +

                            "[製品コード]," +
                            "[薬液コード]," +
                            "[薬品コード1]," +
                            "[薬品コード2]," +
                            "[薬品コード3]," +
                            "[薬品コード4]," +
                            "[薬品コード5]," +
                            "[薬品コード6]," +
                            "[薬品コード7]," +
                            "[薬品コード8]," +
                            "[薬品コード9]," +
                            "[薬品コード10]," +
                            "[薬品コード11]," +
                            "[薬品コード12]," +
                            "[薬品コード13]," +
                            "[薬品コード14]," +
                            "[薬品コード15]," +
                            "[薬品コード16]," +
                            "[薬品コード17]," +
                            "[薬品コード18]," +
                            "[薬品コード19]," +
                            "[薬品コード20]," +
                            "[薬品コード21]," +
                            "[薬品コード22]," +
                            "[薬品コード23]," +
                            "[薬品コード24]," +
                            "[薬品コード25]," +
                            "[薬品コード26]," +
                            "[薬品コード27]," +
                            "[薬品コード28]," +
                            "[薬品コード29]," +
                            "[薬品コード30]," +
                            
                            "[入庫コード]," +                          
                            "[出庫コード]," +
                            "[機械コード]," +
                            "[日付]" +
                            ") VALUES ('" +
                            m_UserCode + "','" +
                            f_ItemDataGet("製品名") + "','" +
                            f_ItemDataGet("薬液名") + "','" +
                            f_ItemDataGet("処方番号") + "','" +
                            f_ItemDataGet("薬品名1") + "','" +
                            f_ItemDataGet("薬品名2") + "','" +
                            f_ItemDataGet("薬品名3") + "','" +
                            f_ItemDataGet("薬品名4") + "','" +
                            f_ItemDataGet("薬品名5") + "','" +
                            f_ItemDataGet("薬品名6") + "','" +
                            f_ItemDataGet("薬品名7") + "','" +
                            f_ItemDataGet("薬品名8") + "','" +
                            f_ItemDataGet("薬品名9") + "','" +
                            f_ItemDataGet("薬品名10") + "','" +
                            f_ItemDataGet("薬品名11") + "','" +
                            f_ItemDataGet("薬品名12") + "','" +
                            f_ItemDataGet("薬品名13") + "','" +
                            f_ItemDataGet("薬品名14") + "','" +
                            f_ItemDataGet("薬品名15") + "','" +
                            f_ItemDataGet("薬品名16") + "','" +
                            f_ItemDataGet("薬品名17") + "','" +
                            f_ItemDataGet("薬品名18") + "','" +
                            f_ItemDataGet("薬品名19") + "','" +
                            f_ItemDataGet("薬品名20") + "','" +
                            f_ItemDataGet("薬品名21") + "','" +
                            f_ItemDataGet("薬品名22") + "','" +
                            f_ItemDataGet("薬品名23") + "','" +
                            f_ItemDataGet("薬品名24") + "','" +
                            f_ItemDataGet("薬品名25") + "','" +
                            f_ItemDataGet("薬品名26") + "','" +
                            f_ItemDataGet("薬品名27") + "','" +
                            f_ItemDataGet("薬品名28") + "','" +
                            f_ItemDataGet("薬品名29") + "','" +
                            f_ItemDataGet("薬品名30") + "','" +
                            f_ItemDataGet("作液工程名1") + "','" +
                            f_ItemDataGet("作液工程名2") + "','" +
                            f_ItemDataGet("作液工程名3") + "','" +
                            f_ItemDataGet("作液工程名4") + "','" +
                            f_ItemDataGet("作液工程名5") + "','" +
                            f_ItemDataGet("作液工程名6") + "','" +
                            f_ItemDataGet("作液工程名7") + "','" +
                            f_ItemDataGet("作液工程名8") + "','" +
                            f_ItemDataGet("作液工程名9") + "','" +
                            f_ItemDataGet("作液工程名10") + "','" +
                            f_ItemDataGet("作液工程名11") + "','" +
                            f_ItemDataGet("作液工程名12") + "','" +
                            f_ItemDataGet("作液工程名13") + "','" +
                            f_ItemDataGet("作液工程名14") + "','" +
                            f_ItemDataGet("作液工程名15") + "','" +
                            f_ItemDataGet("作液工程名16") + "','" +
                            f_ItemDataGet("作液工程名17") + "','" +
                            f_ItemDataGet("作液工程名18") + "','" +
                            f_ItemDataGet("作液工程名19") + "','" +
                            f_ItemDataGet("作液工程名20") + "','" +
                            f_ItemDataGet("作液工程名21") + "','" +
                            f_ItemDataGet("作液工程名22") + "','" +
                            f_ItemDataGet("作液工程名23") + "','" +
                            f_ItemDataGet("作液工程名24") + "','" +
                            f_ItemDataGet("作液工程名25") + "','" +
                            f_ItemDataGet("作液工程名26") + "','" +
                            f_ItemDataGet("作液工程名27") + "','" +
                            f_ItemDataGet("作液工程名28") + "','" +
                            f_ItemDataGet("作液工程名29") + "','" +
                            f_ItemDataGet("作液工程名30") + "','" +
                            f_ItemDataGet("配合率1") + "','" +
                            f_ItemDataGet("配合率2") + "','" +
                            f_ItemDataGet("配合率3") + "','" +
                            f_ItemDataGet("配合率4") + "','" +
                            f_ItemDataGet("配合率5") + "','" +
                            f_ItemDataGet("配合率6") + "','" +
                            f_ItemDataGet("配合率7") + "','" +
                            f_ItemDataGet("配合率8") + "','" +
                            f_ItemDataGet("配合率9") + "','" +
                            f_ItemDataGet("配合率10") + "','" +
                            f_ItemDataGet("配合率11") + "','" +
                            f_ItemDataGet("配合率12") + "','" +
                            f_ItemDataGet("配合率13") + "','" +
                            f_ItemDataGet("配合率14") + "','" +
                            f_ItemDataGet("配合率15") + "','" +
                            f_ItemDataGet("配合率16") + "','" +
                            f_ItemDataGet("配合率17") + "','" +
                            f_ItemDataGet("配合率18") + "','" +
                            f_ItemDataGet("配合率19") + "','" +
                            f_ItemDataGet("配合率20") + "','" +
                            f_ItemDataGet("配合率21") + "','" +
                            f_ItemDataGet("配合率22") + "','" +
                            f_ItemDataGet("配合率23") + "','" +
                            f_ItemDataGet("配合率24") + "','" +
                            f_ItemDataGet("配合率25") + "','" +
                            f_ItemDataGet("配合率26") + "','" +
                            f_ItemDataGet("配合率27") + "','" +
                            f_ItemDataGet("配合率28") + "','" +
                            f_ItemDataGet("配合率29") + "','" +
                            f_ItemDataGet("配合率30") + "','" +
                            f_ItemDataGet("備考1") + "','" +
                            f_ItemDataGet("備考2") + "','" +
                            f_ItemDataGet("規格1") + "','" +
                            f_ItemDataGet("規格2") + "','" +
                            f_ItemDataGet("規格3") + "','" +
                            f_ItemDataGet("規格4") + "','" +
                            f_ItemDataGet("規格5") + "','" +
                            f_ItemDataGet("規格6") + "','" +
                            f_ItemDataGet("規格7") + "','" +
                            f_ItemDataGet("規格8") + "','" +

                            f_ItemDataGet("製品コード") + "','" +
                            f_ItemDataGet("薬液コード") + "','" +
                            f_ItemDataGet("薬品コード1") + "','" +
                            f_ItemDataGet("薬品コード2") + "','" +
                            f_ItemDataGet("薬品コード3") + "','" +
                            f_ItemDataGet("薬品コード4") + "','" +
                            f_ItemDataGet("薬品コード5") + "','" +
                            f_ItemDataGet("薬品コード6") + "','" +
                            f_ItemDataGet("薬品コード7") + "','" +
                            f_ItemDataGet("薬品コード8") + "','" +
                            f_ItemDataGet("薬品コード9") + "','" +
                            f_ItemDataGet("薬品コード10") + "','" +
                            f_ItemDataGet("薬品コード11") + "','" +
                            f_ItemDataGet("薬品コード12") + "','" +
                            f_ItemDataGet("薬品コード13") + "','" +
                            f_ItemDataGet("薬品コード14") + "','" +
                            f_ItemDataGet("薬品コード15") + "','" +
                            f_ItemDataGet("薬品コード16") + "','" +
                            f_ItemDataGet("薬品コード17") + "','" +
                            f_ItemDataGet("薬品コード18") + "','" +
                            f_ItemDataGet("薬品コード19") + "','" +
                            f_ItemDataGet("薬品コード20") + "','" +
                            f_ItemDataGet("薬品コード21") + "','" +
                            f_ItemDataGet("薬品コード22") + "','" +
                            f_ItemDataGet("薬品コード23") + "','" +
                            f_ItemDataGet("薬品コード24") + "','" +
                            f_ItemDataGet("薬品コード25") + "','" +
                            f_ItemDataGet("薬品コード26") + "','" +
                            f_ItemDataGet("薬品コード27") + "','" +
                            f_ItemDataGet("薬品コード28") + "','" +
                            f_ItemDataGet("薬品コード29") + "','" +
                            f_ItemDataGet("薬品コード30") + "','" +

                            f_ItemDataGet("入庫コード") + "','" +
                            f_ItemDataGet("出庫コード") + "','" +
                            f_ItemDataGet("機械コード") + "','" +
                            f_ItemDataGet("日付") + "'" +
                            ")";

                        

#if DEBUG
                        f_LogWrite(3, retstr + command.CommandText);
#endif
                        command.ExecuteNonQuery();

                    }
                    catch (SqlException ex)
                    {
                        retstr = "エラーが発生しました。)";
                        f_LogWrite(3, retstr + "[" + ex.Message + "]");　//これでいいか確認を行うこと。
                        

                        switch (ex.Number)
                        {
                            case 2627:
                            case -2:
                                retstr = "二重登録エラーが発生しました。(" + function_name + ")";

                                break;
                            default:
                                retstr = "データベースエラーが発生しました。1(" + ex.Message + "[" + function_name + "])";
                                break;

                        }
                        f_LogWrite(3, retstr + "[" + ex.Message + "]");
                        //トランザクションのロールバック
                        transaction.Rollback();
                        return retstr;
                    }
                    catch (Exception ex)
                    {
                        retstr = "例外エラーが発生しました。1(" + ex.Message + "[" + function_name + "])";
                        return retstr;
                        //throw;
                    }

                    // トランザクションのコミット
                    transaction.Commit();

                    conn.Close();
                    conn.Dispose();
                }
                catch (Exception ex)
                {
                    retstr = "例外エラーが発生しました。2(" + function_name + ")";
                    f_LogWrite(3, retstr + "[" + ex.Message + "]" + "[" + ex.StackTrace + "]");
                    return retstr;
                }
                finally
                {

                }
            }

            return retstr;
        }
        //------------------------------------------------------------------------------
        // 共通関数
        // ここをカスタマイズしてください。
        //------------------------------------------------------------------------------

        public int BeforeUpdate(ref string retstr)
        {
            //ここに必要な処理を追加してください。
            if (m_FormCD == 工程内検査表)
            {
                //QRコード設定
                setQRCodeToKouteniKensa();
                retstr = "登録ボタン(工程内検査表でBeforeUpdate)が押下されました。";
                f_LogWrite(3, retstr);
            }
            if (m_FormCD == 作液指示書_配合記録)
            {
                //QRコード設定
                setQRCodeToSakueki();
            }
            if (m_FormCD == 品管用_作液指示書)
            {
                    retstr = "登録ボタン(BeforeUpdate)が押下されました。";
                    f_LogWrite(3, retstr);
                    retstr = "";
            }

            return 0;
            	// 返値は目的に応じてそれぞれ意味づけしてください
        }


        public int AfterUpdate(ref string retstr)
        {
            //ここに必要な処理を追加してください。
            if (m_FormCD == 工程内検査表)
            {               
                //工程内検査表DB登録処理
                //string renban = "00000000001";

                //入力チェック
                //retstr = checkItemKoteinaikensa(ref renban);
                //if (!retstr.Equals(string.Empty))
                //{
                //    return 11;
                //}

                //登録処理
                //retstr = registKoteinaikensa(renban);
                //if (retstr != string.Empty)
                //{
                //    f_LogWrite(3, retstr);
                //    return 11;
                //}

                

            }
            if (m_FormCD == 品管用_作液指示書)
            {
                string seihinmei = f_ItemDataGet("製品名");
                string yakuekimei = f_ItemDataGet("薬液名");

                f_LogWrite(3, "登録start[" +seihinmei + "]");

                //作液指示書品管用DB登録処理
                //入力チェック
                retstr = checkItemSakuekiHinkan(ref seihinmei, yakuekimei);
                if (!retstr.Equals(string.Empty))
                {
                    return 11;
                }

                //入力必須項目が入力されていた場合は登録処理に進む
                else
                {
                    retstr = registSakuekihaigoHinkan(seihinmei, yakuekimei);
                    if (retstr != string.Empty)
                    {
                        f_LogWrite(3, retstr);
                        return 11;
                    }
                }
            }

            // add start CSV作成
            if (m_FormCD == 作液指示書_配合記録　|| m_FormCD == 工程内検査表)
            {
                // CSV作成
                retstr = makeCSV();
                if (retstr != string.Empty)
                {
                    f_LogWrite(3, retstr);
                    return 0;
                }
            }
            // add end CSV作成 

            return 0;   // 返値は目的に応じてそれぞれ意味づけしてください
        }
        /// <summary>
        ///新規フラグ設定処理
        ///引数：ChangeFlg　新規フラグにセットする値
        /// </summary>
        public void NewDataFlgChange(Boolean ChangeFlg)
        {

            if (GlobalVar.NewDataFlg.ContainsKey(m_UserCode))
            {
                GlobalVar.NewDataFlg[m_UserCode] = ChangeFlg;
            }
            else
            {
                GlobalVar.NewDataFlg.Add(m_UserCode, ChangeFlg);
            }
        }
        /// <summary>
        ///申請フラグ設定処理
        ///引数：S_ChangeFlg　申請フラグにセットする値
        /// </summary>
        public void ShinSeiFlgChange(string S_ChangeFlg)
        {

            if (GlobalVar.ShinSeiFlg.ContainsKey(m_UserCode))
            {
                GlobalVar.ShinSeiFlg[m_UserCode] = S_ChangeFlg;
            }
            else
            {
                GlobalVar.ShinSeiFlg.Add(m_UserCode, S_ChangeFlg);
            }
        }
        /// <summary>
        ///工程内検査画面にQRコードをセットする
        ///引数：S_ChangeFlg　申請フラグにセットする値
        /// </summary>
        public void setQRCodeToKouteniKensa()
        {
            string tgensi = "原紙ロットNo1";
            string tfilm = "フィルムロットNo1";
            string tflap = "フラップロットNo1";
            int start = 1;
            string hattoriConfig = m_hattoriConfig + "_" + m_FormCD + ".config";

            f_LogWrite(3, "hattoriCongig=" +  hattoriConfig + "]");

            XmlDocument control = new XmlDocument();
            control.Load(hattoriConfig);
            // 原紙QRコードの最終位置
            m_gensiQREnd = int.Parse(control.GetElementsByTagName("gensi")[0].InnerText);
            // フィルムQRコードの最終位置
            m_filmQREnd = int.Parse(control.GetElementsByTagName("film")[0].InnerText);
            // フラップQRコードの最終位置
            m_flapQREnd = int.Parse(control.GetElementsByTagName("flap")[0].InnerText);

            start = getGensiQRStart(tgensi, m_gensiQREnd);
            File.WriteAllText(@"C:\FormPatData1\temp\QRstart.txt", "(" + start + ")");
            GetQRCode("原紙QR", tgensi, start, m_gensiQREnd);           // QRコードを原紙に取込

            start = getGensiQRStart(tfilm, m_filmQREnd);
            GetQRCode("フィルムQR", tfilm, start, m_filmQREnd);        // QRコードをフィルムに取込

            start = getGensiQRStart(tflap, m_flapQREnd);
            GetQRCode("フラップQR", tflap, start, m_flapQREnd);        // QRコードをフラップに取込
        }

        /// <summary>
        ///工程内検査画面にQRコードをセットする
        ///引数：S_ChangeFlg　申請フラグにセットする値
        /// </summary>
        public void setQRCodeToSakueki()
        {
            string retstr = string.Empty;
            string rotNo = "";
            string rotName = "";
            string QRcode = "";
            bool chgFlg = false;  // QR読み込みフラグ tru:読み込み

            EWebForm xform = new EWebForm();
            xform.m_FileEncoding = m_eForm.m_FileEncoding;  //
            xform.m_ReadXmlDataMode = 2;                    //Xmlデータ読込方式(2:DOM使用)
            xform.XmlDataArray_Load(m_eForm.m_DataFile);    //Xmlデータ読込

            // QRコード読取対応
            for (int i = 1; i <= 30; i++)
            {
                rotName = "薬液ロットNo" + i;              
                rotNo = Strings.StrConv(f_ItemDataGet(rotName), VbStrConv.Narrow);
                if (rotNo.Contains(","))
                {
                    //カンマ区切りで分割    
                    string[] qrcode = rotNo.Split(',');

                    retstr = qrcode[0];
                    if (m_Mode == 2)
                    {
                        xform.XmlDataArray_Change(rotName, qrcode[0], 1, 1);       // Action時更新用

                        chgFlg = true;
                    }
                    else
                    {
                        m_eForm.ChangeControlData(rotName, qrcode[0]);             // BeforeUpadate時更新用
                    }
                }
                else
                {
                    // Action時更新用
                    if (m_Mode == 2)
                    {
                        // 現在の値を読む(手で消したとき対応)
                        QRcode = f_ItemDataGet(rotName);
                        if (QRcode.Equals(""))
                        {
                            //File.WriteAllText(@"C:\FormPatData1\temp\QR_" + rotName + ".txt", "(" + m_Mode + ")(" + rotName + ")[" + " " + "]");
                            xform.XmlDataArray_Change(rotName, " ", 1, 1);       //
                            chgFlg = true;
                        }
                    }
                }
            }
            if (chgFlg)
            {
                xform.XmlDataArray_Save();      //xxxxx_eFormLibTemp.Xmlに保存
            }
        }

        public int Action(int mode, ref string retstr)
        {
            m_Mode = mode;

            //ここに必要な処理を追加してください。
            retstr = "";
            string c_value = string.Empty;
            string entry_day = string.Empty;
            string gouki = string.Empty;
            string tyoku = string.Empty;

            //retstr = "Actionに入りました。mode[" + mode + "]";
            //f_LogWrite(3, retstr);

            switch (mode)
            {
#region //計算
                case 2:
                    if (m_FormCD == 工程内検査表)
                    {
                        setQRCodeToKouteniKensa();
                    }
                    if (m_FormCD == 作液指示書_配合記録)
                    {
                        setQRCodeToSakueki();
                    }

                    break;
#endregion

#region //修正のPage_Load時
                case 4:
                    
                    break;
#endregion


#region //参照作成時Page_Load時
                case 6:
                    
                    break;
#endregion

#region //新規のPage_Load時
                case 1:                                       
                    

                    break;
#endregion

#region //7:削除前処理
                case 7:
                    if (m_FormCD == 工程内検査表)
                    {
                        // 工程内検査表 削除
                        f_delete_Koteinaikensa();

                    }
                    break;
#endregion
#region // 31:削除後処理
                case 31:
                    retstr = "削除されました。mode[" + mode + "]";
                    f_LogWrite(3, retstr);
                    break;
#endregion
#region //計算
                case 15:
                   
                    break;
#endregion
#region // 140(決裁)のとき
                case 18:
                    
                    break;
#endregion

#region // 1000(登録ボタン押下判定)のとき
                case 1000:
                    if (m_FormCD == 作液指示書_配合記録)
                    {
                        if (f_ItemDataGet("結果4") == "" && f_ItemDataGet("結果5") == "") { break; }//1
                        else if (f_ItemDataGet("結果4") != "" && f_ItemDataGet("結果5") == "")//2
                        {
                            double kekka_pH = double.Parse(f_ItemDataGet("結果4"));
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") != "")
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                if ((kekka_pH < pH_1) || (kekka_pH > pH_2))
                                {   //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "")
                            {                               
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                if (kekka_pH > pH_2)
                                {   //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") == "")
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                if (kekka_pH < pH_1)
                                {   //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else
                            {
                                break;
                            }
                        }
                        else if (f_ItemDataGet("結果4") == "" && f_ItemDataGet("結果5") != "")//3
                        {
                            double kekka = double.Parse(f_ItemDataGet("結果5"));
                            if (f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") != "")
                            {
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                if ((kekka < hiju_1) || (kekka > hiju_2))
                                {   //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else if (f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")
                            {
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                if (kekka > hiju_2)
                                {   //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else if (f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") == "")
                            {
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                if (kekka < hiju_1)
                                {   //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            else
                            {
                                break;
                            }
                        }
                        else if (f_ItemDataGet("結果4") != "" && f_ItemDataGet("結果5") != "")
                        {
                            double kekka_pH = double.Parse(f_ItemDataGet("結果4"));
                            double kekka = double.Parse(f_ItemDataGet("結果5"));
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") == "")//20
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                if (kekka_pH < pH_1)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")//4
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                
                                if ((kekka_pH < pH_1 || pH_2 < kekka_pH) && hiju_2 >= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if ((kekka_pH >=pH_1 && pH_2 >= kekka_pH) && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if ((kekka_pH < pH_1 || pH_2 < kekka_pH) && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") == "")//5
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
         
                                if (kekka_pH < pH_1 || pH_2 < kekka_pH)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }                     
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") == "")//6
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));                               
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));

                                if (kekka_pH < pH_1 && hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH < pH_1 && hiju_1 <= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH >= pH_1 && hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")
                            {
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));

                                if (kekka_pH > pH_2 && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH > pH_2 && hiju_2 >= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH <= pH_2 && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")//7
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));

                                if (kekka_pH < pH_1 && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH < pH_1 && hiju_2 >= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH >= pH_1 && hiju_2 < kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") == "")//8
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));

                                if ((kekka_pH < pH_1 || pH_2 < kekka_pH)&& hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if ((kekka_pH < pH_1 || pH_2 < kekka_pH) && hiju_1 <= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if ((kekka_pH >= pH_1 && pH_2 >= kekka_pH) && hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") != "")//9
                            {
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));

                                if (kekka_pH < pH_1  && (hiju_2 < kekka || hiju_1 > kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH < pH_1 && (hiju_2 >= kekka && hiju_1 <= kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH >= pH_1 && (hiju_2 < kekka || hiju_1 > kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") == "")//10
                            {
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                if (kekka_pH > pH_2)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") == "")//11
                            {
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));

                                if (kekka_pH > pH_2 && hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH > pH_2 && hiju_1 <= kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH <= pH_2 && hiju_1 > kekka)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") != "")//12
                            {                               
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));

                                if (kekka_pH > pH_2 && (hiju_1 > kekka || hiju_2 < kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH > pH_2 && (hiju_1 <= kekka && hiju_2 >= kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka_pH <= pH_2 && (hiju_1 > kekka || hiju_2 < kekka))
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") == "")//13
                            {
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                if (kekka < hiju_1)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") != "")//14
                            {
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                if (kekka < hiju_1 || kekka > hiju_2)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")//15
                            {                             
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                if (kekka > hiju_2)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else { break; }
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") == "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") == "")//16
                            {
                                break;
                            }
                            if (f_ItemDataGet("pH1") == "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") == "" && f_ItemDataGet("比重2") != "")//17
                            {
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                                if(kekka > hiju_2 && kekka_pH > pH_2)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka <= hiju_2 && kekka_pH > pH_2)
                                {
                                    //ポップアップを表示
                                    retstr = "ｐHが基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else if (kekka > hiju_2 && kekka_pH <= pH_2)
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }
                                else
                                {
                                    break;
                                }
                            }                           
                            if (f_ItemDataGet("pH1") != "" && f_ItemDataGet("pH2") != "" && f_ItemDataGet("比重1") != "" && f_ItemDataGet("比重2") != "")
                            {
                                //画面データ取得
                                double hiju_1 = double.Parse(f_ItemDataGet("比重1"));
                                double hiju_2 = double.Parse(f_ItemDataGet("比重2"));
                               
                                double pH_1 = double.Parse(f_ItemDataGet("pH1"));
                                double pH_2 = double.Parse(f_ItemDataGet("pH2"));
                               

                                if ((kekka_pH < pH_1) || (kekka_pH > pH_2))
                                {
                                    if ((kekka < hiju_1) || (kekka > hiju_2))
                                    {
                                        //ポップアップを表示
                                        retstr = "ｐHと比重が基準値から外れています。登録しますか？";
                                        return 20;
                                    }
                                    else
                                    {
                                        //ポップアップを表示
                                        retstr = "ｐHが基準値から外れています。登録しますか？";
                                        return 20;
                                    }
                                }
                                else if ((pH_1 <= kekka_pH && kekka_pH <= pH_2) && !(hiju_1 <= kekka && kekka <= hiju_2))
                                {
                                    //ポップアップを表示
                                    retstr = "比重が基準値から外れています。登録しますか？";
                                    return 20;
                                }                              
                            }
                        }
                       

                        else
                        {
                            break;
                        }
                    }
                    break;
#endregion

                
            }
            return 0;   // 返値は目的に応じてそれぞれ意味づけしてください
        }

        /// <summary>
        /// 項目名称の設定
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename"></param>
        /// <param name="rtnMsg"></param>
        /// <returns></returns>
        private void f_ItemNameSet()
        {
            // 指示書No
            for (int i = 1; i <= 12; i++)
            {
                //GlobalVar.title_directionNo[i]  = "指示書No" + i.ToString();
                //GlobalVar.title_product_kbn[i]  = "製品区分" + i.ToString();
                //GlobalVar.title_product_code[i] = "商品コード" + i.ToString();
            }
            for(int i = 1; i<=30; i++)
            {
                //GlobalVar.title_hinmoku_no[i] = "品目番号_" + i.ToString();
                //GlobalVar.title_comment[i] = "コメント_" + i.ToString();
            }
            
        }

        /// <summary>
        /// ファイルオープン
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename"></param>
        /// <param name="rtnMsg"></param>
        /// <returns></returns>
        private StreamWriter f_OpenFile(string path, string filename, ref string rtnMsg)
        {
            StreamWriter writer = null;
            string msg = "";
            try
            {
                //StreamWriterを使ってファイルを開いたときは、別のプロセスから書き込みはできない
                //別プロセスが同様の処理をしている場合は、例外が発生する
                // 出力先フォルダがない場合は作成する
                if (Directory.Exists(path) == false)
                {
                    Directory.CreateDirectory(path);
                }
                writer = new StreamWriter(path + "\\" + filename, true, Encoding.GetEncoding("shift_jis"));
                return writer;
            }
            catch (IOException ex)
            {
                msg = "ファイルが開けません。時間をおいて作業してください。";
                f_LogWrite(2, msg + " [" + ex.Message + "]");
                rtnMsg = msg;
                return writer;

            }
            catch (Exception ex)
            {
                msg = "例外エラーが発生しました。";
                f_LogWrite(3, msg + "[" + ex.Message + "]" + "[" + ex.StackTrace + "]");
                rtnMsg = msg;
                return writer;

            }
        }

        

        

        

        /// <summary>
        /// 工程内検査　削除
        /// </summary>
        /// <returns>項目定義の有無(true:定義あり false:定義なし</returns>
            private string f_delete_Koteinaikensa()
        {
            string retstr = "";
            //string function_name = "f_delete_Dailyproductioreport";

            //control.configよりデータベースの接続先を取得する
            XmlDocument control = new XmlDocument();
            control.Load(m_ControlConfig);
            string strdatabase = control.GetElementsByTagName(m_DatabaseName)[0].InnerText;

            using (var conn = new SqlConnection(strdatabase))
            {
                try
                {
                    // DB接続
                    var command = conn.CreateCommand();
                    conn.Open();

                    string renban = f_ItemDataGet("renban");
                    if (renban == string.Empty)
                    {
                        retstr ="連番が発番されていません。";
                        return retstr;
                    }

                    //トランザクションの開始
                    //SqlTransaction transaction = conn.BeginTransaction();

                    //command.Connection = conn;
                    //command.Transaction = transaction;

                    //連番で削除
                    command.CommandText = "DELETE FROM 服部製紙様用_工程内検査表 WHERE 連番 = '" + renban + "'";
                                                     
#if DEBUG
                    f_LogWrite(3, retstr + command.CommandText);
#endif
                    command.ExecuteNonQuery();

                    command.CommandText = "DELETE FROM 服部製紙様用_工程内検査表 WHERE 連番 = '" + renban + "'";
#if DEBUG
                    f_LogWrite(3, retstr + command.CommandText);
#endif
                    command.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    //retstr = "データベースエラーが発生しました。2(" + function_name + ")";
                    retstr = "データベースエラーが発生しました。)";
                    f_LogWrite(3, retstr + "[" + ex.Message + "]");
                    return retstr;

                }
                catch (Exception ex)
                {
                    //retstr = "例外エラーが発生しました。2(" + function_name + ")";
                    retstr = "例外エラーが発生しました。)";
                    f_LogWrite(3, retstr + "[" + ex.Message + "]" + "[" + ex.StackTrace + "]");
                    return retstr;
                }
                finally
                {

                }
                return retstr;
            }
        }


        private string nultozero(string value)
        {
            if (value.Equals(""))
            {
                return "0";
            }
            else
            {
                return value;
            }
        }

        
                        
        /// <summary>
        /// 画面データ取得
        /// </summary>
        /// <param name="ItemName"項目名></param>
        /// <param name="rtnStr">指定された項目のデータ文字列</param>
        /// <returns>項目定義の有無(true:定義あり false:定義なし</returns>
        private string f_ItemDataGet(string ItemName)
        {
            // オブジェクト取得
            string str = "";
            if (m_Mode == 2)//計算時はArrayより取得(Arrayには予めセット済み)
            {
                int i_idx = Array.IndexOf(m_eForm.IDFormArray, ItemName);
                if (i_idx >= 0)
                {
                    str = m_eForm.TextArray[i_idx];
                }
                return str;
            }

            object obj = null;
            bool bl = m_eForm.GetControleObject(ref obj, ItemName);     //画面よりｺﾝﾄﾛｰﾙを取得
            if (bl == false)
            {
                return "";
            }

            string tp = obj.GetType().ToString();
            string c_Type = obj.GetType().ToString();

            //ｺﾝﾄﾛｰﾙの種類により入力値を取得
            switch (c_Type)
            {
                case "System.Web.UI.WebControls.Label":     //ラベル 
                    System.Web.UI.WebControls.Label o_lbl = (System.Web.UI.WebControls.Label)obj;
                    if (o_lbl.Controls.Count > 0)
                        str = ((Label)o_lbl.Controls[0]).Text.Trim();
                    else
                        str = o_lbl.Text.Trim();
                    break;

                case "System.Web.UI.WebControls.TextBox":   //テキストボックス
                    System.Web.UI.WebControls.TextBox o_txt = (System.Web.UI.WebControls.TextBox)obj;
                    str = o_txt.Text.Trim();
                    break;

                case "System.Web.UI.WebControls.DropDownList":  //コンボボックス
                    System.Web.UI.WebControls.DropDownList o_lst = (System.Web.UI.WebControls.DropDownList)obj;
                    str = o_lst.SelectedValue.ToString().Trim();
                    break;

                case "System.Web.UI.WebControls.CheckBox":      //チェックボックス
                    System.Web.UI.WebControls.CheckBox o_chk = (System.Web.UI.WebControls.CheckBox)obj;
                    str = o_chk.Checked.ToString();
                    break;

                case "System.Web.UI.WebControls.RadioButton":   //ラジオボタン
                    System.Web.UI.WebControls.RadioButton o_rdo = (System.Web.UI.WebControls.RadioButton)obj;
                    str = o_rdo.Checked.ToString();
                    break;
            }

            return str;
        }

        // ページ数取得
        private int f_PageCountGet()
        {
            return m_eForm.m_PageCount; // ページ数
        }

        // シート数取得
        private int f_SheetCountGet(int PageNo)
        {
            if ((PageNo < 1) || (PageNo > m_eForm.m_PageCount)) return 0;

            try
            {
                int SheetCount = m_eForm.PageFormList[PageNo].m_PageSheetCount;  // シート数
                return SheetCount;
            }
            catch
            {
                return 1;
            }
        }

        

        public static string TrimDoubleQuotationMarks(string target)
        {
            return target.Trim(new char[] { '"' });
        }


        public string getNewestFileName(string folderName, string key)
        {
#if DEBUG
            GlobalVar.strtemp = "debug3 folderName=" + "(" + folderName + ")" + "key=" + "(" + key + ")";
#endif
            // 指定されたフォルダ内のcsvファイル名をすべて取得する
            string[] files = System.IO.Directory.GetFiles(folderName, "*" + key + "*.csv", System.IO.SearchOption.TopDirectoryOnly);
#if DEBUG
            GlobalVar.strtemp = "debug4 folderName=" + "(" + folderName + ")" + "key=" + "(" + key + ")";
#endif

            string newestFileName = string.Empty;
            System.DateTime updateTime = System.DateTime.MinValue;
            foreach (string file in files)
            {
                // それぞれのファイルの更新日付を取得する
                System.IO.FileInfo fi = new System.IO.FileInfo(file);
                // 更新日付が最新なら更新日付とファイル名を保存する
                if (fi.LastWriteTime > updateTime)
                {
                    updateTime = fi.LastWriteTime;
                    newestFileName = file;
                }
            }
            // ファイル名を返す
            return System.IO.Path.GetFileName(newestFileName);
        }

        

        private void f_LogDelete(string logpath)
        {
            try
            {
                string delDate = DateTime.Now.AddDays(-31).ToString("yyyyMMdd");
                string[] filearray = Directory.GetFiles(logpath, "TOSCO_*.log");
                Array.Sort(filearray);


                foreach (string file in filearray)
                {
                    string filename = Path.GetFileName(file);
                    if (String.Compare(delDate, filename.Substring(6, 8)) > 0)
                    {
                        string delfile = logpath + "\\" + filename;
                        File.Delete(delfile);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception)
            {
                // 削除できない場合はそのまま
            }
        }
        /// <summary>
        /// ログ出力
        /// </summary>
        /// <param name="level">ログレベル(0:DEBUG 1:INFO 2:WARNING 3:ERROR )</param>
        /// <param name="msg">ログメッセージ</param>
        private void f_LogWrite(int level, string msg)
        {
            try
            {
                // ログレベルが設定された最小ログレベルよりも低い場合はログを出力しない
                if (m_MinLogLevel > level)
                {
                    return;
                }

                string logpath = "";


                // C:\\FormPat\\control.configに定義されているパスにログを出力する
                XmlDocument control = new XmlDocument();
                control.Load(m_ControlConfig);

                logpath = control.GetElementsByTagName("addr_abs_temp")[0].InnerText;

                f_LogDelete(logpath);

                string strlevel = "[D]";

                //ファイル名はTOSCO_日付.log
                string Path = logpath + "\\" + "TOSCO_" + DateTime.Now.ToString("yyyyMMdd") + ".log";

                switch (level)
                {
                    case 0:
                        strlevel = "[D]";
                        break;
                    case 1:
                        strlevel = "[I]";
                        break;
                    case 2:
                        strlevel = "[W]";
                        break;
                    case 3:
                        strlevel = "[E]";
                        break;
                }
                string[] strLog = new string[] { DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss:fff ") + " " + m_FormCD + " " + m_Page.Session["UserCode"] + " " + strlevel + " " + msg };
                File.AppendAllLines(Path, strLog, Encoding.GetEncoding("Shift_JIS"));




            }
            catch (Exception)
            {

                //ログに書けない場合もそのまま
            }
        }

        /// <summary>
        /// CSVファイルを作成する
        /// </summary>
        /// <param name="formid">FormID</param>
        public string makeCSV()
        {
            string retStr = string.Empty;
            int iret = 0;

            string hattoriConfig = m_hattoriConfig + "_" + m_FormCD + ".config";

            f_LogWrite(3, "CSV Header Make start["+ m_FormName + "]");


            XmlDocument control = new XmlDocument();
            control.Load(hattoriConfig);
            string headerDir = control.GetElementsByTagName("headerdir")[0].InnerText;
            string csvOutDir = control.GetElementsByTagName("csvoutdir")[0].InnerText;

            string hfileName = headerDir + "\\" + m_FormCD + "_HEAD.csv";
            string vfileName = csvOutDir + "\\" + m_FormCD + "_" +m_UserCode + ".csv";
            DataSet dataSet = new DataSet();
            DataTable table = new DataTable("Table");

            //f_LogWrite(3, "CSV Header Make start2");

            iret =  CsvHToDataTable(table, true, hfileName, ",", false);
            if(iret <0)
            {
                // エラー
                retStr = "CSVヘッダーの読取が出来ませんでした[" + hfileName + "]";
                return retStr;
            }

            // DataSetにDataTableを追加
            dataSet.Tables.Add(table);

            //f_LogWrite(3, "CSV Header Make end[" + hfileName + "]");

            // DataRowクラスを使ってデータを追加
            DataRow dr = table.NewRow();
            foreach (DataColumn column in table.Columns)
            {
                Console.WriteLine(column.ColumnName);
                dr[column.ColumnName] = f_ItemDataGet(f_ItemNameGet(1, 1, column.ColumnName));

                //f_LogWrite(3, "CSV[" + dr[column.ColumnName] + "]");
            }

            dataSet.Tables["Table"].Rows.Add(dr);

            if (System.IO.File.Exists(vfileName))
            {
                // CSV追記出力
                iret = DataTableToCsv(table, vfileName, false);

                f_LogWrite(3, "追記出力 CSVFile[" + vfileName + "]");
            }
            else
            {
                iret = DataTableToCsv(table, vfileName, true);

                f_LogWrite(3, "新規出力 CSVFile[" + vfileName + "]");
            }
            if(iret < 0)
            {
                retStr = "CSV出力エラー[" + vfileName + "]"; 
            }

            //f_LogWrite(3, "CSV Make end");

            return retStr;
        }
        /// <summary>
        /// DataTableの内容をCSVファイルに保存する
        /// </summary>
        /// <param name="dt">CSVに変換するDataTable</param>
        /// <param name="csvPath">保存先のCSVファイルのパス</param>
        /// <param name="writeHeader">ヘッダを書き込む時はtrue。</param>
        public int DataTableToCsv(DataTable dt, string csvPath, bool writeHeader)
        {
            //CSVファイルに書き込むときに使うEncoding
            System.Text.Encoding enc =
                System.Text.Encoding.GetEncoding("Shift_JIS");
            System.IO.StreamWriter sr;

            try
            {
                if (writeHeader)
                {
                    //書き込むファイルを開く（上書き）
                   sr = new System.IO.StreamWriter(csvPath, false, enc);
                }
                else
                {
                    //書き込むファイルを開く(追記）
                    sr = new System.IO.StreamWriter(csvPath, true, enc);
                }

                int colCount = dt.Columns.Count;
                int lastColIndex = colCount - 1;

                //ヘッダを書き込む
                if (writeHeader)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        //ヘッダの取得
                        string field = dt.Columns[i].Caption;
                        //"で囲む
                        field = EncloseDoubleQuotesIfNeed(field);
                        //フィールドを書き込む
                        sr.Write(field);
                        //カンマを書き込む
                        if (lastColIndex > i)
                        {
                            sr.Write(',');
                        }
                    }
                    //改行する
                    sr.Write("\r\n");
                }

                //レコードを書き込む
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        //フィールドの取得
                        string field = row[i].ToString();
                        //"で囲む
                        field = EncloseDoubleQuotesIfNeed(field);
                        //フィールドを書き込む
                        sr.Write(field);
                        //カンマを書き込む
                        if (lastColIndex > i)
                        {
                            sr.Write(',');
                        }
                    }
                    //改行する
                    sr.Write("\r\n");
                }
                //閉じる
                sr.Close();
            }
            catch (Exception ex)
            {
                string retstr = "";
                retstr = "ConvertDataTableToCsv(1) 例外エラーが発生しました。";
                f_LogWrite(3, retstr + "[" + ex.Message + "]" + "[" + ex.StackTrace + "]");

                return -1;
            }
            finally
            {
                string retstr = "CSVデータ出力完了:" + csvPath;
                f_LogWrite(3, retstr);
            }
            return 0;
        }
        /// <summary>
        /// CSVファイルのHeader内容をDataTableに登録する
        /// </summary>
        /// <param name="dt">データを入れるDataTable</param>
        /// <param name="hasHeader">CSVの一行目がカラム名かどうか</param>
        /// <param name="fileName">ファイル名</param>
        /// <param name="separator">カラムを分けている文字(,など)</param>
        /// <param name="quote">カラムを囲んでいる文字("など)</param>
        public  int CsvHToDataTable(DataTable dt, bool hasHeader, string fileName, string separator, bool quote)
        {
            try
            {
                //CSVを便利に読み込んでくれるTextFieldParserを使います。
                TextFieldParser parser = new TextFieldParser(fileName, Encoding.GetEncoding("shift_jis"));
                //これは可変長のフィールドでフィールドの区切りのマーカーが使われている場合です。
                //フィールドが固定長の場合は
                //parser.TextFieldType = FieldType.FixedWidth;
                parser.TextFieldType = FieldType.Delimited;
                //区切り文字を設定します。
                parser.SetDelimiters(separator);
                //クォーテーションがあるかどうか。
                //但しダブルクォーテーションにしか対応していません。シングルクォーテーションは認識しません。
                parser.HasFieldsEnclosedInQuotes = quote;
                string[] data;
                //ここのif文では、DataTableに必要なカラムを追加するために最初に1行だけ読み込んでいます。
                //データがあるか確認します。
                if (!parser.EndOfData)
                {
                    //CSVファイルから1行読み取ります。
                    data = parser.ReadFields();
                    //カラムの数を取得します。
                    int cols = data.Length;
                    if (hasHeader)
                    {
                        for (int i = 0; i < cols; i++)
                        {
                            dt.Columns.Add(new DataColumn(data[i]));
                        }
                    }
                    else
                    {
                        for (int i = 0; i < cols; i++)
                        {
                            //カラム名にダミーを設定します。
                            dt.Columns.Add(new DataColumn());
                        }
                        //DataTableに追加するための新規行を取得します。
                        DataRow row = dt.NewRow();
                        for (int i = 0; i < cols; i++)
                        {
                            //カラムの数だけデータをうつします。
                            row[i] = data[i];
                        }
                        //DataTableに追加します。
                        dt.Rows.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                f_LogWrite(3, "CSVヘッダーファイル読取エラー[" + ex.Message + "]");
                return -1;
            }
            return 0;

        }

        /// <summary>
        /// CSVファイルの内容をDataTableに保存する
        /// </summary>
        /// <param name="dt">データを入れるDataTable</param>
        /// <param name="hasHeader">CSVの一行目がカラム名かどうか</param>
        /// <param name="fileName">ファイル名</param>
        /// <param name="separator">カラムを分けている文字(,など)</param>
        /// <param name="quote">カラムを囲んでいる文字("など)</param>
        public void CsvToDataTable(DataTable dt, bool hasHeader, string fileName, string separator, bool quote)
        {
            //CSVを便利に読み込んでくれるTextFieldParserを使います。
            TextFieldParser parser = new TextFieldParser(fileName, Encoding.GetEncoding("shift_jis"));
            //これは可変長のフィールドでフィールドの区切りのマーカーが使われている場合です。
            //フィールドが固定長の場合は
            //parser.TextFieldType = FieldType.FixedWidth;
            parser.TextFieldType = FieldType.Delimited;
            //区切り文字を設定します。
            parser.SetDelimiters(separator);
            //クォーテーションがあるかどうか。
            //但しダブルクォーテーションにしか対応していません。シングルクォーテーションは認識しません。
            parser.HasFieldsEnclosedInQuotes = quote;
            string[] data;
            //ここのif文では、DataTableに必要なカラムを追加するために最初に1行だけ読み込んでいます。
            //データがあるか確認します。
            if (!parser.EndOfData)
            {
                //CSVファイルから1行読み取ります。
                data = parser.ReadFields();
                //カラムの数を取得します。
                int cols = data.Length;
                if (hasHeader)
                {
                    for (int i = 0; i < cols; i++)
                    {
                        dt.Columns.Add(new DataColumn(data[i]));
                    }
                }
                else
                {
                    for (int i = 0; i < cols; i++)
                    {
                        //カラム名にダミーを設定します。
                        dt.Columns.Add(new DataColumn());
                    }
                    //DataTableに追加するための新規行を取得します。
                    DataRow row = dt.NewRow();
                    for (int i = 0; i < cols; i++)
                    {
                        //カラムの数だけデータをうつします。
                        row[i] = data[i];
                    }
                    //DataTableに追加します。
                    dt.Rows.Add(row);
                }
            }
            //ここのループがCSVを読み込むメインの処理です。
            //内容は先ほどとほとんど一緒です。
            while (!parser.EndOfData)
            {
                data = parser.ReadFields();
                DataRow row = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row[i] = data[i];
                }
                dt.Rows.Add(row);
            }
        }

        /// <summary>
        /// 必要ならば、文字列をダブルクォートで囲む
        /// </summary>
        private string EncloseDoubleQuotesIfNeed(string field)
        {
            if (NeedEncloseDoubleQuotes(field))
            {
                return EncloseDoubleQuotes(field);
            }
            return field;
        }

        /// <summary>
        /// 文字列をダブルクォートで囲む
        /// </summary>
        private string EncloseDoubleQuotes(string field)
        {
            if (field.IndexOf('"') > -1)
            {
                //"を""とする
                field = field.Replace("\"", "\"\"");
            }
            return "\"" + field + "\"";
        }

        /// <summary>
        /// 文字列をダブルクォートで囲む必要があるか調べる
        /// </summary>
        private bool NeedEncloseDoubleQuotes(string field)
        {
            return field.IndexOf('"') > -1 ||
                field.IndexOf(',') > -1 ||
                field.IndexOf('\r') > -1 ||
                field.IndexOf('\n') > -1 ||
                field.StartsWith(" ") ||
                field.StartsWith("\t") ||
                field.EndsWith(" ") ||
                field.EndsWith("\t");
        }
        // 複数シート用項目名取得
        public string f_ItemNameGet(int PageNo, int SheetNo, string ItemName)
        {
            if (ItemName.Trim() == "") return "";

            if (PageNo < 1) return ItemName;

            if (SheetNo <= 1) return ItemName;

            if (PageNo > m_eForm.m_PageCount) return "";

            try
            {
                int SheetCount = m_eForm.PageFormList[PageNo].m_PageSheetCount;  // シート数
                if (SheetNo > SheetCount) return "";
            }
            catch
            {
                return "";
            }

            string ItemHead = HEADCHAR.Replace("+n1", PageNo.ToString()).Replace("+n2", SheetNo.ToString());  // 複数シート用ヘッダ文字

            return ItemHead + ItemName;
        }
    }

}

