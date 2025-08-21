import { Component, OnInit } from '@angular/core';
import { AdminService, Administrator } from '../../../services/admin/admin.service';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-spradm-admin',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './spradm-admin.component.html',
  styleUrl: './spradm-admin.component.css'
})
export class SpradmAdminComponent implements OnInit {
  usuarios: any[] = [];

  constructor(private adminService: AdminService) { }

  ngOnInit(): void {
    this.adminService.obtenerUsuarios().subscribe({
      next: (data) => this.usuarios = data,
      error: (e) => console.error(e)
    });
  }

}
