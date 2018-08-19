/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
import { Component } from '@angular/core';
import { PlayersComponent } from './components/players/players.component';

import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
	selector: 'sheetjs',
	template: `
	
	<div class="inputDiv">
		<input class="fileInput" type="file"  (change)="onFileChange($event)" multiple="false" />
		
	</div>
	
		
	
	<app-players [players]="data" 
				(messageEvent)="receiveMessage($event)"
				[export]="myExportFunc"></app-players>

	`,
	styleUrls: ['./components/players/players.component.css'],
})

export class SheetJSComponent {
	data: AOA = [ [1, 2], [3, 4] ];
	winnerArr: [String, number, number, number][];
	

	receiveMessage($event) {
	  this.winnerArr = $event
	  //console.log("receiveMessage", this.winnerArr);
	}
	wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'SheetJS.xlsx';

	onFileChange(evt: any) {
		/* wire up file reader */
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1) throw new Error('Cannot use multiple files');
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
			/* read workbook */
			const bstr: string = e.target.result;
			const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

			/* grab first sheet */
			const wsname: string = wb.SheetNames[0];
			const ws: XLSX.WorkSheet = wb.Sheets[wsname];

			/* save data */
			this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
			
			//console.log(this.data);
		};
		reader.readAsBinaryString(target.files[0]);
	}

	export(): void {
		/* generate worksheet */
		
		const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.removeDuplicates(this.winnerArr)); 
		
		/* generate workbook and add the worksheet */
		const wb: XLSX.WorkBook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, 'winners');

		/* save to file */
		console.log("before save ");
		//XLSX.writeFile(wb, this.fileName);
		///
		/* from app code, require('electron').remote calls back to main process */
		let dialog = require('electron').remote.dialog;

		/* show a file-open dialog and read the first selected file */
		//let o = dialog.showOpenDialog({ properties: ['openFile'] });
		//let workbook = XLSX.readFile(o[0]);

		/* show a file-save dialog and write the workbook */
		let o = dialog.showSaveDialog({ 
			filters: [{
				name: '',
				extensions: ['xlsx']
			}] });
		XLSX.writeFile(wb, o );
		///
		console.log("after save ");
	}

	get myExportFunc() {
        return this.export.bind(this);
	}
	
	removeDuplicates(arr) {
		let unique_array = []
		for(let i = 0;i < arr.length; i++){
			if(unique_array.indexOf(arr[i]) == -1){
				unique_array.push(arr[i])
			}
		}
		return unique_array
	}
}
