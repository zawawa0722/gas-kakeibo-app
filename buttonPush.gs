// ボタンフラグ取得
function onButtonClick() {
  // 実行前に確認メッセージを表示
  let result = Browser.msgBox("確認", "家計簿出力処理を実行してもよろしいですか？", Browser.Buttons.YES_NO);

  // YESボタンが押された場合のみスクリプトを実行
  if (result == "yes") {
    let buttonPressed = true;
    importCsv(buttonPressed);
  } else {
    // NOボタンが押された場合は実行しない
    Logger.log("家計簿出力処理は実行されませんでした");
  }
}
