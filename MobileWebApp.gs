function doGet() {
  return HtmlService.createHtmlOutputFromFile("mobile/MobileDashboard")
    .setTitle("家計簿")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover");
}
