import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Observable } from 'rxjs';
import MD5 from 'crypto-js/md5';

@Injectable({
  providedIn: 'root'
})

export class AuthService {

  private apiUrl = 'http://localhost:8000/api/login_plain/'; // URL del endpoint de autenticación

  constructor(private http: HttpClient) { }

  login(username: string, password: string): Observable<any> {
    const hashedPassword = MD5(password).toString(); // Hashea la contraseña con MD5

    const headers = new HttpHeaders(
      {
        'Content-Type': 'application/json'
      
      }
    );

  // Enviamos username y password hasheada
  return this.http.post<any>(this.apiUrl, { username, password: hashedPassword }, { headers });
  }

  logout() {
    localStorage.removeItem('token'); // Elimina el token del almacenamiento local
  }

  guardarToken(token: string) {
    localStorage.setItem('token', token); // Guarda el token en el almacenamiento local
  }

  obtenerToken() {
    return localStorage.getItem('token'); // Obtiene el token del almacenamiento local
  }

}
