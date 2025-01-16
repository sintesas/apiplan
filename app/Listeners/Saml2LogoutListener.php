<?php

namespace App\Listeners;

use Aacotroneo\Saml2\Events\Saml2LogoutEvent;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Queue\InteractsWithQueue;

class Saml2LogoutListener
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
    public function handle(Saml2LogoutEvent $event): void
    {
        session()->forget('nameId');
        session()->forget('sessionIndex');
        session()->flush();
    }
}
