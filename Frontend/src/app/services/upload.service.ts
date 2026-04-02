import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface UploadResponse {
  message: string;
  output_file: string;
  download_url: string;
}

@Injectable({
  providedIn: 'root'
})
export class UploadService {
  private baseUrl = 'https://vitals-backend-rs89.onrender.com';

  constructor(private http: HttpClient) {}

  uploadFile(file: File): Observable<UploadResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<UploadResponse>(`${this.baseUrl}/upload`, formData);
  }
}