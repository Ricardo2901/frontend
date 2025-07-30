import { Component } from '@angular/core';
import { Title } from '@angular/platform-browser';

@Component({
  selector: 'app-admin-help',
  standalone: true,
  imports: [],
  templateUrl: './admin-help.component.html',
  styleUrl: './admin-help.component.css'
})
export class AdminHelpComponent {
  constructor(private titleService: Title) {
      this.setTituloPagina('Centro de Ayuda | Administrador')
    }
  
    setTituloPagina(titulo: string) {{
      this.titleService.setTitle(titulo);
    }}
}
