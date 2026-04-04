import { Component, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { finalize } from 'rxjs/operators';
import { UploadService, UploadResponse } from '../services/upload.service';
import { ChangeDetectorRef } from '@angular/core';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './upload.component.html',
  styleUrls: ['./upload.component.css']
})
export class UploadComponent {
  @ViewChild('fileInput') fileInput!: ElementRef<HTMLInputElement>;

  selectedFile: File | null = null;
  isLoading = false;

  message = '';
  error = '';
  downloadUrl = '';
  outputFile = '';

  toastMessage = '';
  toastType: 'success' | 'error' | '' = '';

  constructor(
    private uploadService: UploadService,
    private cdr: ChangeDetectorRef
  ) {}

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;

    if (input.files && input.files.length > 0) {
      this.selectedFile = input.files[0];
      this.clearMessages();
    } else {
      this.selectedFile = null;
    }
  }

  onUpload(): void {
    if (!this.selectedFile) {
      this.showToast('Please select a file.', 'error');
      return;
    }

    if (this.isLoading) {
      return;
    }

    this.isLoading = true;
    this.clearMessages();

    this.uploadService
      .uploadFile(this.selectedFile)
      .pipe(
        finalize(() => {
          this.isLoading = false;
          this.cdr.detectChanges();
          console.log('finalize called, loading reset');
        })
      )
      .subscribe({
        next: (response: UploadResponse) => {
          console.log('Upload success:', response);

          this.message = response.message || 'File processed successfully.';
          this.outputFile = response.output_file || '';
          this.downloadUrl = response.download_url || '';

          if (this.downloadUrl && this.outputFile) {
            this.downloadFile(this.downloadUrl, this.outputFile);
            this.showToast(`Downloaded: ${this.outputFile}`, 'success');
          } else {
            this.showToast('File processed but download link is missing.', 'error');
          }

          this.resetFileInput();
        },
        error: (err) => {
          console.log('Upload error:', err);
          this.error = err?.error?.error || err?.message || 'Upload failed.';
          this.showToast(this.error, 'error');
          this.resetFileInput();
        }
      });
  }

  downloadFile(url: string, filename?: string): void {
    const link = document.createElement('a');
    link.href = url;
    link.download = filename || 'output.docx';
    link.target = '_self';

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  downloadAgain(event?: Event): void {
    if (event) {
      event.preventDefault();
    }

    if (this.downloadUrl && this.outputFile) {
      this.downloadFile(this.downloadUrl, this.outputFile);
    } else {
      this.showToast('No file available to download again.', 'error');
    }
  }

  showToast(message: string, type: 'success' | 'error'): void {
    this.toastMessage = message;
    this.toastType = type;

    setTimeout(() => {
      this.toastMessage = '';
      this.toastType = '';
    }, 3000);
  }

  clearMessages(): void {
    this.message = '';
    this.error = '';
    this.downloadUrl = '';
    this.outputFile = '';
  }

  resetFileInput(): void {
    this.selectedFile = null;

    if (this.fileInput?.nativeElement) {
      this.fileInput.nativeElement.value = '';
    }
  }

  resetForm(): void {
    this.clearMessages();
    this.resetFileInput();
    this.cdr.detectChanges();
    this.isLoading = false;
  }
}