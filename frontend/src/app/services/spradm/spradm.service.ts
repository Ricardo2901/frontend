import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface SuperAdministrador {
  id: number,
  username: string,
  name: string,
  email: string,
  email_verified_at: string,
  password: string,
  created_at: string,
  updated_at: string,
  last_login: string,
  is_active: number,
  type_user: string,
}

@Injectable({
  providedIn: 'root'
})
export class SpradmService {
  // Cambia la URL según tu configuración del backend
  // Asegúrate de que el backend esté corriendo en el puerto 8000 o en el de tu preferencia
  private apiUrl = 'http://localhost:8000/api/root-benutzername/';

  constructor(private http: HttpClient) { }

  obtenerUsuarios(): Observable<SuperAdministrador[]> {
    return this.http.get<SuperAdministrador[]>(this.apiUrl);
    }
}
