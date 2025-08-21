import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SpradmService, SuperAdministrador } from '../../../services/spradm/spradm.service';

@Component({
  selector: 'app-spradm-spradm',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './spradm-spradm.component.html',
  styleUrl: './spradm-spradm.component.css'
})
export class SpradmSpradmComponent implements OnInit {
  usuarios: any[] = [];
  
  constructor(private spradmService: SpradmService) { }

  ngOnInit(): void {
    this.spradmService.obtenerUsuarios().subscribe({
      next: (data) => this.usuarios = data,
      error: (e) => console.error(e)
    });
  }
}
