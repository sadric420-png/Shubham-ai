import { Component, signal, inject, computed, ChangeDetectionStrategy } from '@angular/core';
import { CommonModule, DatePipe } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormGroup, Validators, FormArray } from '@angular/forms';
import { ExcelService } from './services/excel.service';

interface ChatMessage {
  sender: 'ai' | 'user';
  text: string;
  timestamp: Date;
}

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: [], 
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, ReactiveFormsModule]
})
export class AppComponent {
  private excelService = inject(ExcelService);
  private fb = inject(FormBuilder);

  // Workflow State
  currentStep = signal<number>(1);
  isLoading = signal<boolean>(false);
  
  // Data Stores
  masterData = signal<any[]>([]);
  secondData = signal<any[]>([]);
  templateHeaders = signal<string[]>([]);
  finalProcessedData = signal<any[]>([]);
  
  // Missing Data Management
  missingParties = signal<string[]>([]);
  missingDataForm: FormGroup;

  // Chat/Interaction
  messages = signal<ChatMessage[]>([
    {
      sender: 'ai',
      text: "Namaste! Main aapka Route Manager AI hoon. Kripya apni Master File upload karein taaki hum process shuru kar sakein.",
      timestamp: new Date()
    }
  ]);

  constructor() {
    this.missingDataForm = this.fb.group({
      entries: this.fb.array([])
    });
  }

  get missingEntries() {
    return this.missingDataForm.get('entries') as FormArray;
  }

  addMessage(sender: 'ai' | 'user', text: string) {
    this.messages.update(msgs => [...msgs, { sender, text, timestamp: new Date() }]);
  }

  // Step 1: Upload Master
  async onMasterFileUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.isLoading.set(true);
    const file = input.files[0];
    
    try {
      const data = await this.excelService.readExcel(file);
      this.masterData.set(data);
      this.addMessage('user', `Uploaded Master File: ${file.name}`);
      this.addMessage('ai', "Master File mil gayi hai. Ab kripya apni Second File (Current Data) upload karein jisme aaj ka naya route data hai.");
      this.currentStep.set(2);
    } catch (e) {
      this.addMessage('ai', "Error reading Master File. Kripya valid Excel file upload karein.");
      console.error(e);
    } finally {
      this.isLoading.set(false);
      input.value = ''; // Reset input
    }
  }

  // Step 2: Upload Second File
  async onSecondFileUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.isLoading.set(true);
    const file = input.files[0];

    try {
      const data = await this.excelService.readExcel(file);
      this.secondData.set(data);
      this.addMessage('user', `Uploaded Second File: ${file.name}`);
      this.verifyDataGap();
    } catch (e) {
      this.addMessage('ai', "Error reading Second File. Kripya dubara koshish karein.");
    } finally {
      this.isLoading.set(false);
      input.value = '';
    }
  }

  // Step 3: Verification & Gap Analysis
  verifyDataGap() {
    const master = this.masterData();
    const second = this.secondData();
    
    const getPartyName = (row: any) => {
      if (!row) return '';
      const keys = Object.keys(row);
      const nameKey = keys.find(k => k.toLowerCase().includes('party') || k.toLowerCase().includes('name'));
      return nameKey ? String(row[nameKey]).trim() : String(Object.values(row)[0]).trim(); 
    };

    const masterNames = new Set(master.map(r => getPartyName(r).toLowerCase()));
    const missing: string[] = [];

    // Check second file rows
    second.forEach(row => {
      const name = getPartyName(row);
      if (name && !masterNames.has(name.toLowerCase())) {
        // Prevent duplicates
        if (!missing.includes(name)) missing.push(name);
      }
    });

    this.missingParties.set(missing);

    if (missing.length > 0) {
      this.addMessage('ai', "Mujhe kuch naye Party Names mile hain jo Master Database mein nahi hain. Kripya inki details (Party Name, Phone Number, Address) provide karein.");
      this.prepareMissingDataForm(missing);
      this.currentStep.set(3);
    } else {
      this.addMessage('ai', "Sabhi parties Master Data mein maujood hain. Ab Template Filling shuru karte hain.");
      this.addMessage('ai', "Kripya 'Template File' upload karein.");
      this.currentStep.set(5);
    }
  }

  prepareMissingDataForm(missingNames: string[]) {
    this.missingEntries.clear();
    missingNames.forEach(name => {
      this.missingEntries.push(this.fb.group({
        partyName: [name, Validators.required],
        phone: ['', Validators.required],
        address: ['', Validators.required]
      }));
    });
  }

  // Step 4: Update Virtual Master
  submitMissingData() {
    if (this.missingDataForm.invalid) {
      this.missingDataForm.markAllAsTouched();
      return;
    }

    const newEntries = this.missingDataForm.value.entries;
    
    // Update local master data with new entries
    // We create a structure that mimics likely master columns
    const updatedMaster = [...this.masterData()];
    
    newEntries.forEach((entry: any) => {
       updatedMaster.push({
         'Party Name': entry.partyName,
         'Phone': entry.phone,
         'Address': entry.address,
         // Add fallbacks for other common column names
         'Name': entry.partyName,
         'Contact': entry.phone,
         'Mobile': entry.phone,
         'Location': entry.address
       });
    });

    this.masterData.set(updatedMaster);

    this.addMessage('user', 'Missing details provided.');
    this.addMessage('ai', "Details update ho gayi hain. Ab kripya 'Template File' upload karein.");
    this.currentStep.set(5);
  }

  // Step 5: Template Filling
  async onTemplateFileUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    this.isLoading.set(true);
    const file = input.files[0];

    try {
      const headers = await this.excelService.readHeaders(file);
      this.templateHeaders.set(headers);
      this.addMessage('user', `Uploaded Template: ${file.name}`);
      this.processFinalReport(headers);
    } catch (e) {
      this.addMessage('ai', "Template read karne mein samasya aayi.");
      console.error(e);
    } finally {
      this.isLoading.set(false);
      input.value = '';
    }
  }

  processFinalReport(templateHeaders: string[]) {
    const master = this.masterData();
    const second = this.secondData();
    const finalData: any[] = [];

    // Helper to find data in an object loosely
    const findValue = (obj: any, searchKey: string): any => {
      if (!obj) return null;
      
      // 1. Direct match (Case sensitive)
      if (obj[searchKey] !== undefined) return obj[searchKey];

      // 2. Case insensitive match
      const keyLower = searchKey.toLowerCase();
      const keys = Object.keys(obj);
      const foundKey = keys.find(k => k.toLowerCase() === keyLower);
      if (foundKey) return obj[foundKey];

      // 3. Intelligent Mapping for specific fields if exact match failed
      // Address Mapping
      if (keyLower.includes('address') || keyLower.includes('location')) {
         const addrKey = keys.find(k => k.toLowerCase().includes('address') || k.toLowerCase().includes('location'));
         if (addrKey) return obj[addrKey];
      }
      // Phone Mapping
      if (keyLower.includes('phone') || keyLower.includes('contact') || keyLower.includes('mobile')) {
         const phKey = keys.find(k => k.toLowerCase().includes('phone') || k.toLowerCase().includes('contact') || k.toLowerCase().includes('mobile'));
         if (phKey) return obj[phKey];
      }
      
      return null;
    };

    const getPartyName = (row: any) => {
      if (!row) return '';
      const keys = Object.keys(row);
      const nameKey = keys.find(k => k.toLowerCase().includes('party') || k.toLowerCase().includes('name'));
      return nameKey ? String(row[nameKey]).trim() : String(Object.values(row)[0]).trim(); 
    };

    // Iterate over Second File (The Route) - This dictates the rows in output
    second.forEach(routeRow => {
        const routePartyName = getPartyName(routeRow);
        
        // Find matching Master Row
        const masterRow = master.find(m => getPartyName(m).toLowerCase() === routePartyName.toLowerCase());

        const newRow: any = {};

        // Fill columns based on Template Headers
        templateHeaders.forEach(header => {
            // Priority 1: Check Route File (Second File) for specific order details (e.g. Quantity)
            let val = findValue(routeRow, header);

            // Priority 2: Check Master File for static details (Address, Phone)
            if (val === null || val === '') {
                val = findValue(masterRow, header);
            }

            // Priority 3: If still empty, ensure empty string
            newRow[header] = val !== null ? val : '';
        });

        finalData.push(newRow);
    });

    this.finalProcessedData.set(finalData);
    this.addMessage('ai', "Aapka New Route report taiyar hai. Kripya check kar lein.");
    this.currentStep.set(6);
  }

  downloadReport() {
    const today = new Date().toISOString().slice(0, 10);
    this.excelService.exportExcel(this.finalProcessedData(), `New_Route_Report_${today}.xlsx`);
    this.addMessage('user', 'Report Downloaded');
    this.addMessage('ai', 'Dhanyavaad! Kya aap aur koi route process karna chahenge?');
    // Optional: Reset logic could go here
  }
}
