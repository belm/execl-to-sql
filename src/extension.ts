import * as vscode from 'vscode';
import * as XLSX from 'xlsx';

export function activate(context: vscode.ExtensionContext) {
	let disposable = vscode.commands.registerCommand('extension.excelToSql', () => {
		const editor = vscode.window.activeTextEditor;
		if (!editor) {
			vscode.window.showErrorMessage('请在编辑器中打开一个文档');
			return;
		}

		const text = editor.document.getText();
		if (!text) {
			vscode.window.showErrorMessage('请粘贴Excel数据');
			return;
		}

		const tableName = 'your_table_name'; // 替换为实际的表名
		const workbook = XLSX.read(text, { type: 'string', raw: true });
		const sheet = workbook.Sheets[workbook.SheetNames[0]];
		const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

		if (data.length < 2) {
			vscode.window.showErrorMessage('数据不足以生成SQL');
			return;
		}

		const columns: string[] = data[0] as string[];
		const sqlStatements = (data.slice(1) as any[][]).map((row: any[]) => {
			//const sqlStatements = data.slice(1).map((row: any[]) => {
			const values = row.map((value, index) => {
				return typeof value === 'string' ? `'${value}'` : value;
			}).join(', ');

			return `INSERT INTO ${tableName} (${columns.join(', ')}) VALUES (${values});`;
		}).join('\n');

		editor.edit((editBuilder) => {
			editBuilder.replace(new vscode.Range(editor.document.positionAt(0), editor.document.positionAt(text.length)), sqlStatements);
		});
	});

	context.subscriptions.push(disposable);
}