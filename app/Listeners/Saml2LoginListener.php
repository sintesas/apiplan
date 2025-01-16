<?php

namespace App\Listeners;

use Aacotroneo\Saml2\Events\Saml2LoginEvent;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Queue\InteractsWithQueue;

use App\Models\Usuario;
use App\Models\UsuarioRol;
use App\Models\UsuarioMenu;

class Saml2LoginListener
{
    /**
     * Create the event listener.
     */
    public function __construct()
    {
        //
    }

    /**
     * Handle the event.
     */
    public function handle(Saml2LoginEvent $event): void
    {
        $saml2User = $event->getSaml2User();
        $samlAttributes = $saml2User->getAttributes();
        $userData = array(
            'username' => $saml2User->getUserId(),
            'fullname' => $samlAttributes['FullName'][0],
            'email' => $samlAttributes['Email'][0],
            'assertion' => $saml2User->getRawSamlAssertion(),
            'sessionIndex' => $saml2User->getSessionIndex(),
            'nameId' => $saml2User->getNameId()
        );

        // Verificar si el usuario ya existe y obtener el usuario
        $user = Usuario::where('usuario', $userData['username'])->first();

        // Si el usuario no existe, crea nuevo usuario
        if ($user == null) {
            $user = new Usuario;
            $user->usuario = $userData['username'];
            $user->nombre_completo = $userData['fullname'];
            $user->email = $userData['email'] == null ? null : strtolower($userData['email']);
            $user->usuario_creador = 'admin';
            $user->fecha_creacion = \DB::raw('GETDATE()');
            $user->save();

            if ($user->usuario_id != 0) {
                $urol = new UsuarioRol;
                $urol->usuario_id = $user->usuario_id;
                $urol->rol_id = 4;
                $urol->rol_privilegio_id = 5;
                $urol->activo = 'S';
                $urol->usuario_creador = 'admin';
                $urol->fecha_creacion = \DB::raw('GETDATE()');
                $urol->save();

                $urol = new UsuarioRol;
                $urol->usuario_id = $user->usuario_id;
                $urol->rol_id = 5;
                $urol->rol_privilegio_id = 6;
                $urol->activo = 'S';
                $urol->usuario_creador = 'admin';
                $urol->fecha_creacion = \DB::raw('GETDATE()');
                $urol->save();

                $umenu = new UsuarioMenu;
                $umenu->usuario_id = $user->usuario_id;
                $umenu->menu_id = '1,5,6,10,12';
                $umenu->usuario_creador = 'admin';
                $umenu->fecha_creacion = \DB::raw('GETDATE()');
                $umenu->save();
            }
        }

        session(['nameId' => $userData['nameId']]);
        session(['sessionIndex' => $userData['sessionIndex']]);        

        // login usuario
        \Auth::login($user);
    }
}
