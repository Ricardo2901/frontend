
import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { PrivateFileService, Documento } from '../../../services/private-file/private-file.service';
import { FormGroup, FormsModule, FormControl } from "@angular/forms";
import { ReactiveFormsModule } from '@angular/forms';
import { Router, RouterLink } from '@angular/router';
import * as bootstrap from 'bootstrap';

@Component({
  selector: 'app-spradm-private-files',
  standalone: true,
  imports: [CommonModule, FormsModule, ReactiveFormsModule],
  templateUrl: './spradm-private-files.component.html',
  styleUrl: './spradm-private-files.component.css'
})
export class SpradmPrivateFilesComponent implements OnInit {
  userDocuments: Documento[] = [];

  constructor(private usersService: PrivateFileService) { }

  ngOnInit() {
    this.loadDocuments();
  }

  loadDocuments() {
    this.usersService.getUserDocuments().subscribe({
      next: docs => this.userDocuments = docs,
      error: err => {
        console.error('Error al cargar documentos', err);
        if (err.status === 401) {
          // Redirigir al login
          console.warn('Token inválido o expirado');
        }
      }
    });
  }

  onFileSelected(event: any) {
    const file: File = event.target.files[0];
    if (file) {
      this.usersService.uploadFile(file).subscribe(response => this.loadDocuments());
    }
  }

  deleteDocument(id: number) {
    this.usersService.deleteDocument(id).subscribe(() => this.loadDocuments());
  }
}
