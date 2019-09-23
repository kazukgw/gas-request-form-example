const CONFIG_SHEET_NAME = 'Config';
const LOG_SHEET_NAME = 'Log'
const LOG_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
const LOG_MAX_ROWS = 2000

class AppLogger {
  static log = function(message: any) {
    Logger.log(message);
    LOG_SHEET.appendRow([new Date(), Session.getActiveUser().getEmail(), message]);
    let delteNum: number = LOG_SHEET.getMaxRows() - LOG_MAX_ROWS;
    if (delteNum > 0) {
      LOG_SHEET.deleteRows(1, delteNum);
    }
  }
}

class App {
  public rawFormSheet: RawFormSheet
  public issueSheet: IssueSheet
  public issueRepository: IssueRepository
  public issueDocFolder: IssueDocFolder
}

function initApp(): App {
  AppLogger.log(`Init App`);
  let app = new App();
  const configSheet = new ConfigSheet(CONFIG_SHEET_NAME);
  const config = configSheet.getConfig();

  app.rawFormSheet = new RawFormSheet(config);
  app.issueSheet = new IssueSheet(config);
  app.issueDocFolder = new IssueDocFolder(config);

  app.issueRepository = new IssueRepository(
    app.issueSheet,
    app.rawFormSheet,
    app.issueDocFolder,
    config,
  )
  AppLogger.log(`Init App successfully`);
  return app
}

function onSubmit(event: any) {
  AppLogger.log(`submit new issue: ${JSON.stringify(event)}`)
  let app = initApp();
  // for debug
  //
  // event = {
  //  namedValues: {
  //    "メールアドレス": ["xxxx@gmail.com"],
  //    "タイムスタンプ": ["2019/09/09 01:01:01"]
  //    }
  // };
  //
  let formData = new FormData(event);

  app.issueRepository.newIssueFromFormData(formData);
}

function onChangeIssueStatus(event: any) {
  AppLogger.log(`change sheet: ${JSON.stringify(event)}`);
  let app = initApp();
  let sheet = SpreadsheetApp.getActiveSheet();
  if(sheet.getName() !== app.issueSheet.name) {
    AppLogger.log(`${sheet.getName()} is not ${app.issueSheet.name}}`)
    return;
  }
  let row = event['range']['rowStart'];
  let col = event['range']['columnStart'];
  if(row < 2) {
    return;
  }
  if(col !== app.issueSheet.getStatusColumnNum()) {
    AppLogger.log(`column num ${col} is not ${app.issueSheet.getStatusColumnNum()}}`)
    return;
  }

  let issueIdString = sheet.getRange(row, 1).getValue();
  AppLogger.log(`change issue: id=${issueIdString}`)
  let issueId = IssueId.createByIdString(issueIdString)
  let newStatus = <string>event['value'];
  app.issueRepository.changeIssueStatus(issueId, newStatus);
}
