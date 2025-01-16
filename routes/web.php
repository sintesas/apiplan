<?php

use Illuminate\Support\Facades\Route;

use App\Http\Controllers\Auth\LoginController;
use App\Http\Controllers\PdfController;
use App\Http\Controllers\TestContoller;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', function () {
    return view('index');
});

Route::get('/saml', [LoginController::class, 'saml'])->name('saml');
Route::any('/logout', [LoginController::class, 'logout'])->name('logout');

Route::get('/reporte/perfilCargo/view/{id}', [PdfController::class, 'getInformesPreview']);
Route::get('/reporte/perfilCargo/{id}', [PdfController::class, 'getInformes']);

Route::get('info', function() {
    dd(phpinfo());
});

Route::get('test', [TestContoller::class, 'test']);
