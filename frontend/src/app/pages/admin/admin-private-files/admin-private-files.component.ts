import { Component } from '@angular/core';
import { Title } from '@angular/platform-browser';

@Component({
  selector: 'app-admin-private-files',
  standalone: true,
  imports: [],
  templateUrl: './admin-private-files.component.html',
  styleUrl: './admin-private-files.component.css'
})
export class AdminPrivateFilesComponent {
  constructor(private titleService: Title) {
    this.setTituloPagina('Archivos Privados | Administrador')
  }

  setTituloPagina(titulo: string) {{
    this.titleService.setTitle(titulo);
  }}
}
