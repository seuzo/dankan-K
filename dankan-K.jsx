/*
dankan-K.jsx
選択しているテキストフレームに段間罫をひきます
(c)2009	市川せうぞー

2009-09-09	ver0.1	とりあえず　http://d.hatena.ne.jp/seuzo/20090910/1252519010
2009-09-10	ver0.2	段落罫の長さを本文のブロックよりも罫線を短く（長く）する設定を追加
2009-09-11	ver.0.3	線幅を正しく認識できないバグを修正　http://d.hatena.ne.jp/seuzo/20090911/1252600247
2009-09-11	ver.0.4	生成した段落罫をテキストフレームと一緒にグループ化するかどうか設定できるようにした。
2009-09-13	ver.0.5	GUIをつけて、公開リリース。
*/

////////////////////////////////////////////設定（これらはグローバル変数：ダイアログの初期値を変えたい時はここを変更する）
var G_LINEWIDTH = 0.3; //段間罫の太さ（単位は環境設定に依存）
var G_STROKETYPE = "ベタ";//線種
var G_LINECOLOR = "Black";//罫線のスウォッチ名
var G_OVERLINE = -1;//本文のブロック長よりも、罫線を短く（長く）する。マイナス値の指定で短くなる
var G_GROUP = true;//生成した段落罫をテキストフレームと一緒にグループ化するかどうか。


////////////////////////////////////////////エラー処理 
function myerror(mess) { 
  if (arguments.length > 0) { alert(mess); }
  exit();
}

////////////////////////////////////////////スプレッドの回転角度を調べる （1スプレッド2ページまで対応）
function get_spread_angle(spread_obj) {
	var my_doc, my_old_ruler_origin, my_old_zeroPoint, my_page, my_page_bounds, my_angle, migitoji, i;
	my_doc = app.activeDocument;
	my_old_ruler_origin = false;//
	if (my_doc.viewPreferences.rulerOrigin !== 1380143215) {//not page
		my_old_ruler_origin = my_doc.viewPreferences.rulerOrigin;//current setting
		my_doc.viewPreferences.rulerOrigin = 1380143215;//change ルーラーをページに
	}
	my_old_zeroPoint = false;
	if (my_doc.zeroPoint !== [0, 0]) {
		my_old_zeroPoint = my_doc.zeroPoint;
		my_doc.zeroPoint = [0,0];
	}

	my_page = spread_obj.pages[0];
	my_page_bounds = my_page.bounds;
	migitoji = 0;//右綴じかどうかのカウンター。migitoji===1なら右綴じの見開き
	for (i = 0; i< my_page_bounds.length; i++) {
		if (my_page_bounds[i] === 0) {migitoji++}
	}
	if (migitoji === 1) {//右綴じの見開きは右ページ基点
		my_page = spread_obj.pages[1];
		my_page_bounds = my_page.bounds;
	}

	my_angle = -1;
	if((my_page_bounds[0] === 0) && (my_page_bounds[1] === 0)) {
		my_angle = 0;
	} else if ((my_page_bounds[0] === 0) && (my_page_bounds[3] === 0)) {
		my_angle = 90;
	} else if ((my_page_bounds[2] === 0) && (my_page_bounds[3] === 0)) {
		my_angle = 180;
	} else if ((my_page_bounds[1] === 0) && (my_page_bounds[2] === 0)) {
		my_angle = 270;
	}

	if(my_old_ruler_origin) {my_doc.viewPreferences.rulerOrigin = my_old_ruler_origin}
	if(my_old_zeroPoint) {my_doc.zeroPoint = my_old_zeroPoint}
	return my_angle;
}

////////////////////////////////////////////配列my_arrayの中のmy_itemのindexを返す
function index_ofArray(my_item, my_array) {
	for (var i = 0; i< my_array.length; i++) {
		if (my_item === my_array[i]) {return i;}
	}
	return false;
}

////////////////////////////////////////////ダイアログ（グローバル変数を書き換えるので、返り値なし）
function show_dialog() {
	//前準備
	var my_doc, my_swatches, my_strokeTypes, my_minWidth;
	var my_realEditbox_01, my_dropdown_01, my_dropdown_02, my_realEditbox_02, my_checkbox_01;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	my_doc = app.activeDocument;
	my_strokeTypes = my_doc.strokeStyles.everyItem().name;//ドキュメント中に定義されているすべての線種。名前の配列
	my_swatches = my_doc.swatches.itemByRange(1, -1).name;//ドキュメント中に定義されているすべてのスウォッチ（「なし」以外）。名前の配列
	my_minWidth = 150;//文字列幅
	
	//ダイアログ
	var my_dialog = app.dialogs.add({name:"段間罫を追加", canCancel:true});
	with(my_dialog) {
		with(dialogColumns.add()) {
			// プロンプト
			staticTexts.add({staticLabel:"選択しているテキストフレームに段間罫を追加します。"});
			with (borderPanels.add()) {
				staticTexts.add({staticLabel:"段間罫の太さ：", minWidth:my_minWidth});
				my_realEditbox_01 = realEditboxes.add({editValue:G_LINEWIDTH});
			}
			with(borderPanels.add()){
				staticTexts.add({staticLabel:"段間罫の線種：", minWidth:my_minWidth});
				my_dropdown_01 = dropdowns.add({stringList:my_strokeTypes, selectedIndex:index_ofArray(G_STROKETYPE, my_strokeTypes)});// ポップアップメニュー
			}
			with(borderPanels.add()){
				staticTexts.add({staticLabel:"段間罫の色：", minWidth:my_minWidth});
				my_dropdown_02 = dropdowns.add({stringList:my_swatches, selectedIndex:index_ofArray(G_LINECOLOR, my_swatches)});// ポップアップメニュー
			}
			with (borderPanels.add()) {
				staticTexts.add({staticLabel:"段間罫の調整：", minWidth:my_minWidth});
				my_realEditbox_02 = realEditboxes.add({editValue:G_OVERLINE});
			}
			with (borderPanels.add()) {
				my_checkbox_01 = checkboxControls.add({staticLabel:"段落罫をテキストフレームと一緒にグループ化する", checkedState:G_GROUP});
			}
		}
	}
	
	//ダイアログの表示と値のget
	if (my_dialog.show() == true) {
		G_LINEWIDTH = my_realEditbox_01.editValue;//段間罫の太さ
		my_dropdown_01 = parseInt(my_dropdown_01.selectedIndex);//段間罫の線種
		G_STROKETYPE = my_strokeTypes[my_dropdown_01];//段間罫の線種
		my_dropdown_02 = parseInt(my_dropdown_02.selectedIndex);//段間罫の色
		G_LINECOLOR = my_swatches[my_dropdown_02];//段間罫の色
		G_OVERLINE = my_realEditbox_02.editValue;//段間罫の調整
		G_GROUP = my_checkbox_01.checkedState;//グループ化
		//設定をオブジェクトに変換
		G_STROKETYPE = my_doc.strokeStyles.itemByName(G_STROKETYPE);
		G_LINECOLOR = my_doc.colors.itemByName(G_LINECOLOR);
		//正常にダイアログを片付ける
		my_dialog.destroy();
	} else {
		// ユーザが「キャンセル」をクリックしたので、メモリからダイアログボックスを削除
		my_dialog.destroy();
		exit();
	}
	
	///値のチェック
	if (G_LINEWIDTH <= 0 || G_LINEWIDTH > 100) {
		myerror("段間罫の太さは0以上100以下の範囲内でなければなりません");
	}
	if (G_OVERLINE < -50 || G_OVERLINE > 50) {
		myerror("段間罫の調整は-50以上50以下の範囲内でなければなりません");
	}
}

////////////////////////////////////////////線の色が「なし」のオブジェクトは線幅を0に指定する
function set_stroke2zero(my_obj) {
	var my_doc = app.activeDocument;
	if (my_obj.strokeColor.name === "None") {//線の色が「なし」なら
		my_doc.colors.itemByName("Black");//色を指定して
		my_obj.strokeWeight = 0;//線幅を0に
	}
}

////////////////////////////////////////////オブジェクトの大きさを得る 
function get_bounds(my_obj) {
	var tmp_hash = new Array();
	var my_obj_bounds = my_obj.visibleBounds; //オブジェクトの大きさ（線幅を含む）
	tmp_hash["y1"] = my_obj_bounds[0];
	tmp_hash["x1"] = my_obj_bounds[1];
	tmp_hash["y2"] = my_obj_bounds[2];
	tmp_hash["x2"] = my_obj_bounds[3];
	tmp_hash["w"] = tmp_hash["x2"] - tmp_hash["x1"]; //幅
	tmp_hash["h"] = tmp_hash["y2"] - tmp_hash["y1"]; //高さ
	return tmp_hash; //ハッシュで値を返す
}

////////////////////////////////////////////テキストフレーム設定を得る
function get_textFramePreferences(my_obj) {
	var tmp_hash = new Array();
	tmp_hash["s_weight"] = my_obj.strokeWeight//線の太さ
	tmp_hash["c_count"] = my_obj.textFramePreferences.textColumnCount;//段数
	tmp_hash["c_gutter"] = my_obj.textFramePreferences.textColumnGutter;//段間
	tmp_hash["c_width"] = my_obj.textFramePreferences.textColumnFixedWidth;
	tmp_hash["top"] = my_obj.textFramePreferences.insetSpacing[0];//上マージン
	tmp_hash["left"] = my_obj.textFramePreferences.insetSpacing[1];//左マージン
	tmp_hash["bottom"] = my_obj.textFramePreferences.insetSpacing[2];//左マージン
	tmp_hash["right"] = my_obj.textFramePreferences.insetSpacing[3];//右マージン
	return tmp_hash; //ハッシュで値を返す
}

////////////////////////////////////////////線を引く 
function draw_line(x1, x2, y1, y2, my_lineWidth, my_strokeType, my_lineColor){
	var my_line=app.activeWindow.activeSpread.rectangles.add();
	my_line.paths[0].entirePath = [[x1, y1], [x2, y2]];
	my_line.paths[0].pathType = PathType.OPEN_PATH;
	my_line.strokeWeight = my_lineWidth;
	my_line.strokeType = my_strokeType;
	my_line.strokeColor = my_lineColor;
	my_line.fillColor = "None";
	return my_line;
}


////////////////////////////////////////////以下実行ルーチン
function main(){
var my_doc, my_spread, my_ruler_origin,i, ii, tmp_obj, tmp_obj_bounds, tmp_obj_pref,  tmp_groupItems, f_layerLock, tmp_line, x1, x2, y1, y2; 
if (app.documents.length === 0) {myerror("ドキュメントが開かれていません")}
my_doc = app.activeDocument;
if (my_doc.selection.length === 0) {myerror("テキストフレームを選択してください")}

//スプレッドが回転していたら、エラーで中止。ver6.0以上
if (app.scriptPreferences.version <= 6.0) {
	my_spread = app.layoutWindows[0].activeSpread;
	if (get_spread_angle(my_spread) !== 0) {myerror("スプレッドが回転しています。元に戻してから実行してください。")};
}

//ページルーラーの開始位置が「スプレッド」以外になっていたら、「スプレッド」に一時的に変更
my_ruler_origin = false;//初期値はfalse
if (my_doc.viewPreferences.rulerOrigin !== 1380143983) {
	my_ruler_origin = my_doc.viewPreferences.rulerOrigin;//現在の設定を保存
	my_doc.viewPreferences.rulerOrigin = 1380143983;//「スプレッド」に一時的に変更
}
my_doc.zeroPoint = [0, 0];//ルーラー原点

//ダイアログの表示
show_dialog();

for (i = 0; i < my_doc.selection.length; i++) {
	 tmp_obj = my_doc.selection[i];
	if(tmp_obj instanceof TextFrame) {//テキストフレームなら以下を実行
		set_stroke2zero(tmp_obj);//線の色が「なし」のオブジェクトは線幅を0に指定する
		tmp_obj_bounds = get_bounds(tmp_obj);//オブジェクトの大きさを得る
		tmp_obj_pref = get_textFramePreferences(tmp_obj);//テキストフレーム設定を得る
		tmp_groupItems = new Array();//グループアイテムの初期化
		f_layerLock = tmp_obj.itemLayer.locked;//レイヤーがロックされているかどうか
		tmp_obj.itemLayer.locked = false;//レイヤーのロック解除
		
		for (ii = 1; ii < tmp_obj_pref["c_count"]; ii++) {
			if (tmp_obj.parentStory.storyPreferences.storyOrientation === StoryHorizontalOrVertical.HORIZONTAL) {//横組み
				x1 = tmp_obj_bounds["x1"] + tmp_obj_pref["s_weight"] + tmp_obj_pref["left"] + (tmp_obj_pref["c_width"] * ii) + (tmp_obj_pref["c_gutter"] * ii - tmp_obj_pref["c_gutter"] / 2);
				y1 = tmp_obj_bounds["y1"] + tmp_obj_pref["s_weight"] + tmp_obj_pref["top"] - G_OVERLINE;
				y2 = tmp_obj_bounds["y2"] - tmp_obj_pref["s_weight"] - tmp_obj_pref["bottom"] + G_OVERLINE;
				tmp_line = draw_line(x1, x1, y1, y2, G_LINEWIDTH, G_STROKETYPE, G_LINECOLOR);
			} else if (tmp_obj.parentStory.storyPreferences.storyOrientation === StoryHorizontalOrVertical.VERTICAL) {//縦組み
				y1 = tmp_obj_bounds["y1"] + tmp_obj_pref["s_weight"] + tmp_obj_pref["top"] + (tmp_obj_pref["c_width"] * ii) + (tmp_obj_pref["c_gutter"] * ii - tmp_obj_pref["c_gutter"] / 2);
				x1 = tmp_obj_bounds["x1"] + tmp_obj_pref["s_weight"] + tmp_obj_pref["left"] - G_OVERLINE;
				x2 = tmp_obj_bounds["x2"] - tmp_obj_pref["s_weight"] - tmp_obj_pref["right"] + G_OVERLINE;
				tmp_line = draw_line(x1, x2, y1, y1, G_LINEWIDTH, G_STROKETYPE, G_LINECOLOR);
			}//if
			tmp_groupItems.push(tmp_line);//段間罫をグループの一味に加える
		}//for
		
		if (G_GROUP && (tmp_groupItems.length !== 0) ) {//グループ化指定があれば、グループ化を実行
			if (tmp_obj.parent instanceof Group ) {//すでにグループのひとつだった
				tmp_groupItems.push(tmp_obj.parent);//テキストフレームを含むグループをグループの一味に加える
			} else {
				tmp_groupItems.push(tmp_obj);//テキストフレームをグループの一味に加える
			}
			if (tmp_obj.locked) {
				tmp_obj.locked = false;ロックが掛かっていればロックを外す
				tmp_groupItems = my_spread.groups.add(tmp_groupItems, tmp_obj.itemLayer);//グループ化
				tmp_groupItems.locked = true;
			} else {
				tmp_groupItems = my_spread.groups.add(tmp_groupItems, tmp_obj.itemLayer);//グループ化
			}
		}
		tmp_obj.itemLayer.locked = f_layerLock;//ロックの復帰
	}//if
}//for

//ページルーラー設定の復帰
if (my_ruler_origin) {
	my_doc.viewPreferences.rulerOrigin = my_ruler_origin;
}
}//function

main();