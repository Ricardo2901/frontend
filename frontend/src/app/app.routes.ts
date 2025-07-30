import { Routes } from '@angular/router';
import { UsuariosComponent } from './components/usuarios/usuarios.component';

/* Para Super Usuarios */

/* Para Administradores */
import { AdminComponent } from './pages/admin/admin.component';
import { AdminHomeComponent } from './pages/admin/admin-home/admin-home.component';
import { AdminAboutComponent } from './pages/admin/admin-about/admin-about.component';
import { AdminAdminsComponent } from './pages/admin/admin-admins/admin-admins.component';
import { AdminUsersComponent } from './pages/admin/admin-users/admin-users.component';
import { AdminHelpComponent } from './pages/admin/admin-help/admin-help.component';
import { AdminProfileComponent } from './pages/admin/admin-profile/admin-profile.component';
import { AdminProjectComponent } from './pages/admin/admin-project/admin-project.component';
import { AdminPrivateFilesComponent } from './pages/admin/admin-private-files/admin-private-files.component';

/* Para Usuarios */
import { UserAboutComponent } from './pages/user/user-about/user-about.component';
import { UserHelpComponent } from './pages/user/user-help/user-help.component';
import { UserHomeComponent } from './pages/user/user-home/user-home.component';
import { UserProfileComponent } from './pages/user/user-profile/user-profile.component';
import { UserProjectComponent } from './pages/user/user-project/user-project.component';
import { UserComponent } from './pages/user/user.component';


export const routes: Routes = [
  { path: 'usuarios', component: UsuariosComponent },
  { path: '', redirectTo: '/usuarios', pathMatch: 'full' }, // redirige al inicio

  /* Rutas para Administradores */
  { path: 'mayma/auth/spradm', 
    component: AdminComponent,
    children: [
      { path: '', redirectTo: 'home', pathMatch: 'full' },
      { path: 'home', component: AdminHomeComponent },
      { path: 'about', component: AdminAboutComponent },
      { path: 'admins', component: AdminAdminsComponent },
      { path: 'users', component: AdminUsersComponent },
      { path: 'help', component: AdminHelpComponent },
      { path: 'profile', component: AdminProfileComponent },
      { path: 'project', 
        component: AdminProjectComponent,
        children: [
          { path: 'list', component: AdminPrivateFilesComponent },
          { path: 'name', component: AdminPrivateFilesComponent },
        ]
        /*
        children: [
          { path: 'capitulo-1', component: AdminCapitulo1Component },
          { path: 'capitulo-2', component: AdminCapitulo2Component },
          { path: 'capitulo-3', component: AdminCapitulo3Component },
          { path: 'capitulo-4', component: AdminCapitulo4Component },
          { path: 'capitulo-5', component: AdminCapitulo5Component },
          { path: 'capitulo-6', component: AdminCapitulo6Component },
          { path: 'capitulo-7', component: AdminCapitulo7Component },
          { path: 'capitulo-8', component: AdminCapitulo8Component },
          { path: 'capitulo-9', component: AdminCapitulo9Component },
          { path: 'capitulo-10', component: AdminCapitulo10Component },
          { path: 'capitulo-11', component: AdminCapitulo11Component },
          { path: 'capitulo-12', component: AdminCapitulo12Component },
          { path: 'capitulo-13', component: AdminCapitulo13Component },
          { path: 'capitulo-14', component: AdminCapitulo14Component },
          { path: 'capitulo-15', component: AdminCapitulo15Component },
          { path: 'capitulo-16', component: AdminCapitulo16Component },
          { path: 'capitulo-17', component: AdminCapitulo17Component },
        ]
          */
      },
      { path: 'private-files', component: AdminPrivateFilesComponent}
    ] },

  /* Rutas para Usuarios */
  { path: 'mayma/auth/usr',
    component: UserComponent,
    children: [
      { path: '', redirectTo: 'home', pathMatch: 'full' },  // redirige a home
      { path: 'about', component: UserAboutComponent },
      { path: 'help', component: UserHelpComponent },
      { path: 'home', component: UserHomeComponent },
      { path: 'profile', component: UserProfileComponent },
      { path: 'project', component: UserProjectComponent },
    ]
  }
];
