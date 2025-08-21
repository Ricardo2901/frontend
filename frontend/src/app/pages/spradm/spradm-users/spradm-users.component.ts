import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { UsersService, User } from '../../../services/users/users.service';
import { FormGroup, FormsModule, FormControl } from "@angular/forms";
import { ReactiveFormsModule } from '@angular/forms';
import { Router, RouterLink } from '@angular/router';
import * as bootstrap from 'bootstrap';

@Component({
  selector: 'app-spradm-users',
  standalone: true,
  imports: [CommonModule, FormsModule, ReactiveFormsModule],
  templateUrl: './spradm-users.component.html',
  styleUrl: './spradm-users.component.css'
})
export class SpradmUsersComponent implements OnInit {
  usuarios: User[] = [];

  userForm = new FormGroup({
    username: new FormControl(''),
    name: new FormControl(''),
    email: new FormControl(''),
    password: new FormControl('password'),
    is_active: new FormControl(0),
    type_user: new FormControl('Usuario')
  });

  updateUserForm = new FormGroup({
    id: new FormControl<number | undefined>(undefined),
    username: new FormControl(''),
    name: new FormControl(''),
    email: new FormControl(''),
    password: new FormControl('password'),
    is_active: new FormControl(0),
    type_user: new FormControl('Usuario')
  });

  constructor(private usersService: UsersService) { }

  ngOnInit(): void {
    this.usersService.obtenerUsuarios().subscribe({
      next: (data) => this.usuarios = data,
      error: (e) => console.error(e)
    });
  }

  agregarUsuario() {
    if(this.userForm.valid) {
      this.usersService.agregarUsuario(this.userForm.value).subscribe({
        next: (user) => {
          console.log('Usuario agregado:', user)
        },
        error: (err) => {
          console.error('Error al agregar usuario:', err)
        },
      });
    }
  }

  abrirModal(usuario: User) {
    this.updateUserForm.patchValue({
      id: usuario.id,
      username: usuario.username,
      name: usuario.name,
      email: usuario.email,
      password: '',
      is_active: usuario.is_active,
      type_user: usuario.type_user
    })
  }

  actualizarUsuario() {
    const id = Number(this.updateUserForm.value.id); // convierte null a NaN si no hay valor
    if (!id) return; // seguridad, no enviamos si no hay id válido

    // clonamos el formulario y eliminamos id para enviar solo los datos
    const data = { ...this.updateUserForm.value };
    if (!data.password) delete data.password;

    this.usersService.actualizarUsuario(id, data).subscribe({
      next: (user) => {
        this.ngOnInit()
        console.log('Usuario actualizado:', user);
        // cerrar modal y refrescar lista
      },
      error: (err) => console.error('Error al actualizar usuario:', err)
    });
  }
}
