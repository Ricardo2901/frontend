import { Component } from '@angular/core';
import { Router } from '@angular/router';
import { HttpClient } from '@angular/common/http';
import { AuthService } from '../../services/auth/auth.service';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-main-login',
  standalone: true,
  imports: [
    FormsModule,
  ],
  templateUrl: './main-login.component.html',
  styleUrl: './main-login.component.css'
})

export class MainLoginComponent {
  username: string = '';
  password: string = '';
  errorMessage: string = '';

  constructor(private authService: AuthService, private router: Router) {}

  onSubmit() {
    this.authService.login(this.username, this.password).subscribe(
      { 
        next: (res: any) => {
          this.authService.guardarToken(res.token); // Guarda el token en el almacenamiento local
          this.router.navigate(['/mayma/auth/spradm/home']); // Redirige al usuario a la página de inicio
        },
        error: (err: any) => {
          this.errorMessage = 'Usuario o contraseña incorrectos'; // Muestra un mensaje de error si las credenciales son incorrectas
          console.error(err);
        }
      }
    )
  }
}
