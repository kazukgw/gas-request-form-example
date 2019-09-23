// moment を使えるようにするには以下が必要
// また方などは適宜 import('moment').Moment などして利用する
declare let moment: typeof import('moment');
declare function flat(depth: number | undefined): any;

class FormData {
  public rawEvent: Object
  public values: Object
  public createTime:import('moment').Moment
  public submitterEmail: string
  public summary: string
  public details: string
  public reason: string
  public desiredDeadline: import('moment').Moment
  public severity: string

  constructor(rawEvent: Object) {
    this.rawEvent = rawEvent;
    this.values = rawEvent['namedValues'];

    this.submitterEmail = this.values['メールアドレス'][0];
    this.createTime = moment(this.values['タイムスタンプ'][0], 'YYYY/MM/DD HH:mm:ss', 'ja')
    this.summary = this.values['概要'][0];
    this.details = this.values['詳細'][0];
    this.reason = this.values['理由'][0];
    this.severity = this.values['重要度'][0];
    this.desiredDeadline = moment(this.values['希望期限'][0], 'YYYY/MM/DD', 'ja');
  }
}

class Sheet {
  public name: string
  public headers: Array<string>

  private sheetObj: GoogleAppsScript.Spreadsheet.Sheet

  constructor(name: string) {
    this.sheetObj = SpreadsheetApp.getActive().getSheetByName(name);
    // header の数は最大 20
    let headers = <Array<string>>this.sheetObj.getRange(1, 1, 1, 20).getValues()[0];
    this.headers = [];
    headers.forEach((v)=>{ if(v !== undefined && v !== '') this.headers.push(v) });
  }

  public appendRow(data: Array<string>) {
    this.sheetObj.appendRow(data);
  }

  public lookupRowRangeByColumnValue(headerKey: string, lookupValue: any):
    GoogleAppsScript.Spreadsheet.Range | undefined { let idx: number = this.headers.indexOf(headerKey);
    let rowNum: number = this.sheetObj.getMaxRows();
    let columnValues: Array<Array<any>> = this.sheetObj.getRange(1, idx + 1, rowNum).getValues().flat();
    let rowIdx: number = columnValues.indexOf(lookupValue);
    AppLogger.log(`lookuped rowIdx: ${rowIdx}`)
    if (rowIdx < 0) {
      return undefined
    }
    return this.sheetObj.getRange(rowIdx + 1, 1, 1, this.headers.length + 1);
  }

  public getLastRowRange(): GoogleAppsScript.Spreadsheet.Range {
    return this.sheetObj.getRange(this.sheetObj.getLastRow(), 1, 1, this.headers.length + 1);
  }

  public getKVDataFromA1Notation(a1Notation: string): { [key:string]: string } {
    return this.sheetObj.getRange(a1Notation).getValues().reduce((acc, cur)=>{
      acc[String(cur[0])]  = String(cur[1]);
      return acc
    }, {});
  }
}

class ConfigSheet {
  public name: string
  public sheet: Sheet

  constructor(sheetName: string) {
    this.name = sheetName;
    this.sheet = new Sheet(this.name);
  }

  public getConfig(): Config {
    let configData:{ [key:string]: any } = this.sheet.getKVDataFromA1Notation('A1:B20');
    let config = new Config();

    config.rawFormSheetName = configData["rawFormSheetName"];
    config.issueSheetName = configData["issueSheetName"];
    config.issueDocFolderId = configData["issueDocFolderId"];
    config.issueKey = configData["issueKey"];
    config.defaultEditor = configData["defaultEditor"];
    config.defaultViewer = configData["defaultViewer"];

    return config;
  }
}

class Config {
  public rawFormSheetName: string
  public issueSheetName: string
  public issueDocFolderId: string
  public issueKey: string
  public defaultEditor: string
  public defaultViewer: string
}

class RawFormSheet {
  public name: string
  public sheet: Sheet

  constructor(config: Config) {
    this.name = config.rawFormSheetName;
    this.sheet = new Sheet(this.name);
  }
}

/**
 * IssueSheet のレイアウト, IssueシートからIssueオブジェクトの取得,
 * Issueオブジェクトの永続化に責任を持つ。
 * 基本的に Issue集約としてのみアクセスされることを期待しており,
 * Issue Repository 以外のクラスが直接利用することを想定しない。
 */
class IssueSheet {
  public name: string
  public sheet: Sheet

  constructor(config: Config) {
    this.name = config.issueSheetName;
    this.sheet = new Sheet(this.name);
  }

  public insert(issue: Issue) {
    let data: Array<string> = this.issueToRowData(issue);
    this.sheet.appendRow(data);
  }

  public update(issue: Issue) {
    let range = this.sheet.lookupRowRangeByColumnValue(
      this.getHeaderKeyFromProp('issueId'), issue.issueId);

    if (range === undefined) {
      let err = new Error(`issue(${issue.issueId.toIdString()}) is notfound`)
      AppLogger.log(`Error: ${err}`);
      throw err;
    }

    let dataToSet = [this.issueToRowData(issue)];
    range.setValues(dataToSet);
  }

  public getIssueByIssueId(issueId: IssueId): Issue {
    AppLogger.log(`get issue by issue id: issue id=${issueId}`);
    let range = this.sheet.lookupRowRangeByColumnValue(
      this.getHeaderKeyFromProp('issueId'), issueId.toIdString());

    if (range === undefined) {
      let err = new Error(`issue(${issueId.toIdString()}) is notfound`)
      AppLogger.log(`Error: ${err}`);
      throw err;
    }
    return this.rowDataToIssue(range.getValues().flat());
  }

  public getLatestIssue(): Issue | undefined {
    let range = this.sheet.getLastRowRange();
    if(range.getRow() === 1) {
      return undefined;
    }
    return this.rowDataToIssue(range.getValues().flat());
  }

  public getStatusColumnNum() {
    // TODO: ハードコーディングやめたい
    return this.sheet.headers.indexOf('Status') + 1;
  }

  // TODO: header , sheet layout, issue object の property とのmapping はもっといいかんじに管理したい
  private getHeaderKeyFromProp(prop: string): string | undefined {
    return {
      "issueId": "Issue ID"
    }[prop];
  }

  private rowDataToIssue(rowData: Array<string>): Issue {
    AppLogger.log(`rowData To Issue: ${JSON.stringify(rowData)}`)
    // TODO: index 固定値やめる
    // header と row num の紐付けは自動で解決したい
    // また header の文字列と property name も もっといい感じに解決したい
    let issueIdString = rowData[0];
    let submitter = rowData[1];
    let createTime = moment(rowData[2]);
    let assignee = rowData[3];
    let status = rowData[4];
    let issueDocUrl = rowData[5];

    let issueId = IssueId.createByIdString(issueIdString);
    let issue = new Issue(issueId, submitter, createTime);
    issue.assignee = assignee;
    issue.status = status;
    issue.issueDocUrl = issueDocUrl;
    return issue;
  }

  private issueToRowData(issue: Issue): Array<string>  {
    return [
      issue.issueId.toIdString(),
      issue.submitter,
      issue.createTime.format('YYYY/MM/DD HH:mm:ss'),
      issue.assignee === undefined ? '' : issue.assignee,
      issue.status,
      issue.issueDocUrl,
    ];
  }
}

class IssueId {
  public key: string
  public num: number

  constructor(key: string, num:number) {
    this.key = key;
    this.num = num;
  }

  public toIdString(): string {
    return `${this.key}-${this.num}`;
  }

  public getNextId(): IssueId {
    return new IssueId(this.key, this.num + 1);
  }

  static emptyId(): IssueId { return new IssueId('', 0); }

  static createByIdString(idString: string): IssueId | undefined {
    let values: Array<string> = idString.split('-');
    if(values.length != 2) {
      throw new Error(`invalid id string: ${idString}`)
    }
    return new IssueId(values[0], parseInt(values[1]));
  }
}

class IssueDocFolder {
  public folder: GoogleAppsScript.Drive.Folder;
  constructor(config: Config) {
    this.folder = DriveApp.getFolderById(config.issueDocFolderId);
  }

  addFile(file: GoogleAppsScript.Drive.File) {
    this.folder.addFile(file);
  }
}

class IssueDoc {
  public issue: Issue
  public doc: GoogleAppsScript.Document.Document
  public docFile: GoogleAppsScript.Drive.File

  constructor(issue: Issue) {
    this.issue = issue;
  }

  public saveToFolder(folder: IssueDocFolder) {
    let name = `[${this.issue.status}] ${this.issue.issueId.toIdString()}`;
    this.doc = DocumentApp.create(name);
    this.docFile = DriveApp.getFileById(this.doc.getId());
    folder.addFile(this.docFile);
    DriveApp.getRootFolder().removeFile(this.docFile);
  }

  public addEditor(email: string) {
    this.docFile.addEditor(email)
  }

  public addViewer(email: string) {
    this.docFile.addViewer(email)
  }

  public addContent(formData: FormData) {
    let body = this.doc.getBody();
    let cells = [
      ['作成者', formData.submitterEmail],
      ['作成日時', formData.createTime.format('YYYY/MM/DD HH:mm:ss')],
      ['重要度', formData.severity],
      ['希望期限', formData.desiredDeadline.format('YYYY/MM/DD')],
    ];
    body.appendTable(cells);

    // Append a document header paragraph.
    body.appendParagraph("概要").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(formData.summary);

    body.appendParagraph("詳細").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(formData.details);

    body.appendParagraph("理由").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(formData.reason);

    body.appendParagraph("作業ログ").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  }
}

class Issue {
  public issueId: IssueId
  public submitter: string
  public createTime: import('moment').Moment
  public assignee?: string
  public status: string
  public issueDocUrl?: string

  constructor(issueId: IssueId, submitterEmail: string, createTime: import('moment').Moment) {
    this.issueId = issueId;
    this.submitter = submitterEmail;
    this.createTime = createTime;
    // status の扱いどうするか
    // spread sheet 側で pull down で表示できるようにしたい
    this.status = 'OPEN';
  }
}

/**
 * Issue 集約(IssueDoc含む) の 取得, 生成, 変更, 一貫性に対して責任を持つ
 *
 * - 一貫性とは例えば以下のようなものを指す
 *   - IssueSheet の内容とメモリ上の Issue オブジェクト
 *   - Issue に紐づくIssueDoc
 *   - IssueDoc の内容と Issue オブジェクト
 *
 * - Issueオブジェクトの永続化は IssueSheet クラスが責任をもつ
 */
class IssueRepository {
  public rawFormSheet: RawFormSheet
  public issueSheet: IssueSheet
  public issueDocFolder: IssueDocFolder
  public defaultEditor: string
  public defaultViewer: string
  public issueKey: string

  constructor(
    issueSheet: IssueSheet,
    rawFormSheet: RawFormSheet,
    issueDocFolder: IssueDocFolder,
    config: Config
  ) {
    this.issueSheet = issueSheet;
    this.rawFormSheet = rawFormSheet;
    this.issueDocFolder = issueDocFolder;
    this.defaultEditor = config.defaultEditor;
    this.defaultViewer = config.defaultViewer;
    this.issueKey = config.issueKey;
  }

  public newIssueFromFormData(formData: FormData) { let newId = this.assignNewIssueId();

    let newIssue = new Issue(newId, formData.submitterEmail, formData.createTime);

    let newIssueDoc = new IssueDoc(newIssue);
    newIssueDoc.saveToFolder(this.issueDocFolder);
    newIssueDoc.addEditor(this.defaultEditor);
    newIssueDoc.addEditor(newIssue.submitter);
    newIssueDoc.addContent(formData);
    newIssue.issueDocUrl = newIssueDoc.docFile.getUrl();
    this.issueSheet.insert(newIssue);
  }

  // TODO: もうちょっとちゃんと抽象化したい
  public changeIssueStatus(issueId: IssueId, newStatus: string) {
    AppLogger.log(`change issue status to ${newStatus}`)
    let issue = this.getIssueByIssueId(issueId);
    AppLogger.log(`change issue doc title: docUrl=${issue.issueDocUrl}`);
    let doc = DocumentApp.openByUrl(issue.issueDocUrl);
    let docFile = DriveApp.getFileById(doc.getId());
    let newName = docFile.getName().replace(/^\[[a-zA-Z0-9]+\]/, `[${newStatus}]`);
    docFile.setName(newName);
    AppLogger.log(`change issue doc title successfully (${newName})`)
    // notify to submitter ... ?
  }

  private getIssueByIssueId(issueId: IssueId): Issue {
    return this.issueSheet.getIssueByIssueId(issueId);
  }

  private getMaxIssueId(): IssueId | undefined {
    let latestIssue = this.issueSheet.getLatestIssue();
    if(latestIssue === undefined) {
      return undefined;
    }
    return latestIssue.issueId;
  }

  private assignNewIssueId(): IssueId {
    let maxIssueId = this.getMaxIssueId();
    if (maxIssueId === undefined) {
      return new IssueId(this.issueKey, 1)
    }
    return maxIssueId.getNextId();
  }

}
