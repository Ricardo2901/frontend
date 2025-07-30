import { Component } from '@angular/core';
import { Title } from '@angular/platform-browser';

@Component({
  selector: 'app-admin-home',
  standalone: true,
  imports: [],
  templateUrl: './admin-home.component.html',
  styleUrl: './admin-home.component.css'
})
export class AdminHomeComponent {
  constructor(private titleService: Title) {
    this.setTituloPagina('Inicio | Administrador');
  }

  setTituloPagina(titulo: string) {
    this.titleService.setTitle(titulo);
  }
}
