//設定が必要な項目
var user_name = "ユーザID";
var user_password = "パスワード";
var t_url = "サイトURL";
var t_r_url = "予約ページURL";
var room_name = "部屋名";

//予約の情報
var title = "テスト"
var startyear = "2015";
var startmonth = "12";
var startday = "20";
var starthour = "12";
var startminute = "0";
var endyear = "2015";
var endmonth = "12";
var endday = "20";
var endhour = "13";
var endminute = "0";

//  Internet Explorerオブジェクトを取得(Internet Explorerの起動)
var IEApp = new ActiveXObject( "InternetExplorer.Application" );


//  Internet Explorerアプリケーションを表示
IEApp.Visible = true;


//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );


//　サイトに移動
IEApp.Navigate( t_url );

//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );

// フォームに入力
var TextBox1 = IEApp.document.getElementById( "loginname" );
var TextBox2 = IEApp.document.getElementById( "loginpassword" );
TextBox1.value = user_name;
TextBox2.value = user_password;

// 3秒停止
WScript.Sleep( 3000 );


//  ボタンを押す
var Button = IEApp.document.getElementById( "btn_login" );
Button.click();

//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );

/*-----------------------------------------------------------*/
//　会議室予約ページに移動
IEApp.Navigate( t_r_url );

//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );

//  作成ボタンクリック
var Create =  IEApp.document.getElementsByClassName( "ico-create" ).item(0) ;
Create.click();

//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );

//タイトル
var Title = IEApp.document.getElementsByClassName( "textbox" ).item(0) ;
Title.value = title;

var Startyear = IEApp.document.getElementsByClassName( "textbox" ).item(2) ;
Startyear.value = startyear;

var Startmonth = IEApp.document.getElementsByClassName( "textbox" ).item(3) ;
Startmonth.value = startmonth;

var Startday = IEApp.document.getElementsByClassName( "textbox" ).item(4) ;
Startday.value = startday;

var Starthour = IEApp.document.getElementsByClassName( "textbox" ).item(5) ;
Starthour.value = starthour;

var Startminut = IEApp.document.getElementsByClassName( "textbox" ).item(6) ;
Startminut.value = startminute

//終了時刻
var Endyear = IEApp.document.getElementsByClassName( "textbox" ).item(7) ;
Endyear.value = endyear;

var Endmonth = IEApp.document.getElementsByClassName( "textbox" ).item(8) ;
Endmonth.value = endmonth;

var Endday = IEApp.document.getElementsByClassName( "textbox" ).item(9) ;
Endday.value = endday;

var Endhour = IEApp.document.getElementsByClassName( "textbox" ).item(10) ;
Endhour.value = endhour;

var Endminut = IEApp.document.getElementsByClassName( "textbox" ).item(11) ;
Endminut.value = endminute

//部屋選択
var Room = IEApp.document.getElementsByClassName( "ico-non" ).item(4) ;
Room.click();

// 3秒停止
WScript.Sleep( 3000 );

var SearchForm = IEApp.document.getElementById( "search-form" ).children.item(0).children.item(2).children.item(1).children.item(0);

SearchForm.value = room_name;

var SearchButton = IEApp.document.getElementById( "search-form" ).children.item(0).children.item(3).children.item(0).children.item(0);
SearchButton.click();

// 3秒停止
WScript.Sleep( 3000 );

var RoomButton = IEApp.document.getElementsByClassName( "resrcitem" ).item(0).parentNode;

WScript.Echo( RoomButton.outerHTML　);
RoomButton.dbclick();
// 0.1秒停止
WScript.Sleep( 100 );
RoomButton.click();


/*-----------------------------------------------------------*/
WScript.Echo( SearchForm.outerHTML　);
//  ページの読み込みが終わるまで処理を止める　その後3秒停止
while( IEApp.Busy )
    WScript.Sleep( 500 ); 
WScript.Sleep( 3000 );

//  Internet Explorerを終了
IEApp.Quit();

WScript.Echo( "終了" );

//  オブジェクトを解放
IEApp = null;


