<?php

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;

use App\Http\Controllers\Auth\LoginController;
use App\Http\Controllers\AreaController;
use App\Http\Controllers\CargoController;
use App\Http\Controllers\CuerpoController;
use App\Http\Controllers\EspecialidadController;
use App\Http\Controllers\GradoController;
use App\Http\Controllers\ListaDinamicaController;
use App\Http\Controllers\MenuController;
use App\Http\Controllers\ModuloController;
use App\Http\Controllers\PerfilController;
use App\Http\Controllers\PerfilUsuarioController;
use App\Http\Controllers\RequerimientoController;
use App\Http\Controllers\RequisitoLeyController;
use App\Http\Controllers\RolController;
use App\Http\Controllers\RutaCarreraController;
use App\Http\Controllers\UsuarioController;
use App\Http\Controllers\UsuarioMenuController;
use App\Http\Controllers\ExcelRutaCarreraController;
use App\Http\Controllers\ExcelVariasHojasController;
use App\Http\Controllers\ExcelEscalafonCargoController;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| is assigned the "api" middleware group. Enjoy building your API!
|
*/

// Login
Route::post('login', [LoginController::class, 'login'])->name('login');
// Route::post('login', [UsuarioController::class, 'login']);

Route::group(['middleware' => 'api', 'prefix' => 'auth'], function ($router) {
    // Listas Dinamicas
    Route::get('/listadinamica/getListasDinamicasFull',[ListaDinamicaController::class, 'getListasDinamicasFull']);
    // ---------------------------------------------------------------------------------------------------------------

    // Cargos
    Route::post('/cargo/getCargos', [CargoController::class, 'getCargos']);
    Route::post('/cargo/getCargosId', [CargoController::class, 'getCargosId']);
    Route::post('/cargo/crearCargos', [CargoController::class, 'crearCargos']);
    Route::post('/cargo/actualizarCargos', [CargoController::class, 'actualizarCargos']);
    Route::post('/cargo/eliminarCargos', [CargoController::class, 'eliminarCargos']);
    Route::post('/cargo/getDetalleCargos', [CargoController::class, 'getDetalleCargos']);
    Route::get('/cargo/getCargosFull', [CargoController::class, 'getCargosFull']);
    Route::get('/cargo/getListasNiveles', [CargoController::class, 'getListasNiveles']);
    Route::get('/cargo/consultarRutasPorCargo/{cargo_id}', [CargoController::class, 'consultarRutasPorCargo']);
    Route::get('/cargo/consultarCargosPorCargo/{cargo_previo_id}', [CargoController::class, 'consultarCargosPorCargo']);
    Route::get('/cargo/consultarUbicacionesPorCargo/{cargo_id}', [CargoController::class, 'consultarUbicacionesPorCargo']);

    // Cargos Grado
    Route::post('/cargo/getCargosGradosByCargoId', [CargoController::class, 'getCargosGradosByCargoId']);
    Route::post('/cargo/crearCargosGrados', [CargoController::class, 'crearCargosGrados']);
    Route::post('/cargo/actualizarCargosGrados', [CargoController::class, 'actualizarCargosGrados']);
    Route::post('/cargo/eliminarCargosGradosById', [CargoController::class, 'eliminarCargosGradosById']);

    // Cargos Configuracion
    Route::post('/cargo/getCargosConfiguracionById', [CargoController::class, 'getCargosConfiguracionById']);
    Route::post('/cargo/getCargosConfiguracionByCargoGradoId', [CargoController::class, 'getCargosConfiguracionByCargoGradoId']);
    Route::post('/cargo/crearCargosConfiguracion', [CargoController::class, 'crearCargosConfiguracion']);
    Route::post('/cargo/actualizarCargosConfiguracion', [CargoController::class, 'actualizarCargosConfiguracion']);

    // Cargos Experiencias
    Route::post('/cargo/getCargosExperienciasById', [CargoController::class, 'getCargosExperienciasById']);
    Route::post('/cargo/crearCargosExperiencias', [CargoController::class, 'crearCargosExperiencias']);
    Route::post('/cargo/actualizarCargosExperiencias', [CargoController::class, 'actualizarCargosExperiencias']);
    Route::post('/cargo/eliminarCargosExperienciasId', [CargoController::class, 'eliminarCargosExperienciasId']);

    // Ubicacion de Cargos
    Route::post('/cargo/getUbicacionCargosId', [CargoController::class, 'getUbicacionCargosId']);
    Route::post('/cargo/crearUbicacionCargos', [CargoController::class, 'crearUbicacionCargos']);
    Route::post('/cargo/actualizarUbicacionCargos', [CargoController::class, 'actualizarUbicacionCargos']);
    Route::post('/cargo/eliminarUbicacionCargosId', [CargoController::class, 'eliminarUbicacionCargosId']);
    Route::get('/cargo/exportExcel', [ExcelVariasHojasController::class, 'exportExcel']);

    // Areas
    Route::get('/area/getAreasFull',[AreaController::class, 'getAreasFull']);
    Route::post('/area/getAreas',[AreaController::class, 'getAreas']);
    Route::post('/area/crearAreas',[AreaController::class, 'crearAreas']);
    Route::post('/area/actualizarAreas',[AreaController::class, 'actualizarAreas']);

    // Cuerpos
    Route::get('/cuerpo/getCuerposFull',[CuerpoController::class, 'getCuerposFull']);
    Route::post('/cuerpo/getCuerpos',[CuerpoController::class, 'getCuerpos']);
    Route::post('/cuerpo/crearCuerpos',[CuerpoController::class, 'crearCuerpos']);
    Route::post('/cuerpo/actualizarCuerpos',[CuerpoController::class, 'actualizarCuerpos']);

    // Especialidades
    Route::get('/especialidad/getEspecialidadesFull', [EspecialidadController::class, 'getEspecialidadesFull']);
    Route::post('/especialidad/getEspecialidades',[EspecialidadController::class, 'getEspecialidades']);
    Route::post('/especialidad/crearEspecialidades',[EspecialidadController::class, 'crearEspecialidades']);
    Route::post('/especialidad/actualizarEspecialidades',[EspecialidadController::class, 'actualizarEspecialidades']);

    // Grados
    Route::get('/grado/getGradosFull', [GradoController::class, 'getGradosFull']);
    Route::post('/grado/getGrados', [GradoController::class, 'getGrados']);
    Route::post('/grado/crearGrados', [GradoController::class, 'crearGrados']);
    Route::post('/grado/actualizarGrados', [GradoController::class, 'actualizarGrados']);
    Route::post('/grado/getDetalleGrados', [GradoController::class, 'getDetalleGrados']);

    // Listas Dinamicas
    Route::post('/listadinamica/getNombresListas',[ListaDinamicaController::class, 'getNombresListas']);
    Route::post('/listadinamica/crearNombresListas',[ListaDinamicaController::class, 'crearNombresListas']);
    Route::post('/listadinamica/actualizarNombresListas',[ListaDinamicaController::class, 'actualizarNombresListas']);
    Route::post('/listadinamica/getListasDinamicas',[ListaDinamicaController::class, 'getListasDinamicas']);
    Route::post('/listadinamica/crearListasDinamicas',[ListaDinamicaController::class, 'crearListasDinamicas']);
    Route::post('/listadinamica/actualizarListasDinamicas',[ListaDinamicaController::class, 'actualizarListasDinamicas']);
    Route::get('/listadinamica/getNombresListasFull',[ListaDinamicaController::class, 'getNombresListasFull']);

    // Menu
    Route::post('/menu/getMenus', [MenuController::class, 'getMenus']);
    Route::post('/menu/crearMenus', [MenuController::class, 'crearMenus']);
    Route::post('/menu/actualizarMenus', [MenuController::class, 'actualizarMenus']);

    // Perfiles
    Route::post('/perfil/getPerfiles', [PerfilController::class, 'getPerfiles']);
    Route::post('/perfil/crearPerfiles', [PerfilController::class, 'crearPerfiles']);
    Route::post('/perfil/actualizarPerfiles', [PerfilController::class, 'actualizarPerfiles']);
    Route::get('/perfil/getPerfilesFull', [PerfilController::class, 'getPerfilesFull']);

    // Perfiles-Usuarios
    Route::post('/perfilusuario/getPerfilesUsuarios', [PerfilUsuarioController::class, 'getPerfilesUsuarios']);
    Route::post('/perfilusuario/crearPerfilesUsuario', [PerfilUsuarioController::class, 'crearPerfilesUsuarios']);
    Route::post('/perfilusuario/actualizarPerfilesUsuarios', [PerfilUsuarioController::class, 'actualizarPerfilesUsuarios']);

    // Requisito de Ley
    Route::post('/requisitoley/getRequisitosLey',[RequisitoLeyController::class, 'getRequisitosLey']);
    Route::post('/requisitoley/crearRequisitosLey',[RequisitoLeyController::class, 'crearRequisitosLey']);
    Route::post('/requisitoley/actualizarRequisitosLey',[RequisitoLeyController::class, 'actualizarRequisitosLey']);

    // Requisito de Ley - Especialidades
    Route::post('/requisitoley/getRequisitosEspecialidades',[RequisitoLeyController::class, 'getRequisitosEspecialidades']);
    Route::post('/requisitoley/crearRequisitosEspecialidades',[RequisitoLeyController::class, 'crearRequisitosEspecialidades']);
    Route::post('/requisitoley/actualizarRequisitosEspecialidades',[RequisitoLeyController::class, 'actualizarRequisitosEspecialidades']);

    // Requisito de Ley - Grados
    Route::post('/requisitoley/getRequisitosGrados',[RequisitoLeyController::class, 'getRequisitosGrados']);
    Route::post('/requisitoley/crearRequisitosGrados',[RequisitoLeyController::class, 'crearRequisitosGrados']);
    Route::post('/requisitoley/actualizarRequisitosGrados',[RequisitoLeyController::class, 'actualizarRequisitosGrados']);

    // Ruta de Carrera
    Route::post('/rutacarrera/getRutaCarrera', [RutaCarreraController::class, 'getRutaCarrera']);
    Route::post('/rutacarrera/crearRutaCarrera', [RutaCarreraController::class, 'crearRutaCarrera']);
    Route::post('/rutacarrera/actualizarRutaCarrera', [RutaCarreraController::class, 'actualizarRutaCarrera']);
    Route::post('/rutacarrera/eliminarRutaCarrera', [RutaCarreraController::class, 'eliminarRutaCarrera']);
    Route::post('/rutacarrera/crearLineasCargos', [RutaCarreraController::class, 'crearLineasCargos']);
    Route::post('/rutacarrera/actualizarLineasCargos', [RutaCarreraController::class, 'actualizarLineasCargos']);
    Route::post('/rutacarrera/getLineasCargos', [RutaCarreraController::class, 'getLineasCargos']);
    Route::post('/rutacarrera/getRutas', [RutaCarreraController::class, 'getRutas']);
    Route::post('/rutacarrera/crearRutas', [RutaCarreraController::class, 'crearRutas']);
    Route::post('/rutacarrera/actualizarRutas', [RutaCarreraController::class, 'actualizarRutas']);
    Route::post('/rutacarrera/eliminarRuta', [RutaCarreraController::class, 'eliminarRuta']);
    Route::post('/rutacarrera/getRutasByRutaCarrera', [RutaCarreraController::class, 'getRutasByRutaCarrera']);
    Route::post('/rutacarrera/getCargosByRutas', [RutaCarreraController::class, 'getCargosByRutas']);
    Route::post('/rutacarrera/getGradosByEspecialidad', [RutaCarreraController::class, 'getGradosByEspecialidad']);
    Route::post('/rutacarrera/getGradosDetalleByEspecialidad', [RutaCarreraController::class, 'getGradosDetalleByEspecialidad']);
    Route::post('/rutacarrera/getGradosDetalleCargo', [RutaCarreraController::class, 'getGradosDetalleCargo']);
    Route::post('/rutacarrera/getGradosDetalleRequisitoLey', [RutaCarreraController::class, 'getGradosDetalleRequisitoLey']);
    Route::post('/rutacarrera/getCuerposByCategoria', [RutaCarreraController::class, 'getCuerposByCategoria']);
    Route::post('/rutacarrera/getEspecialidadesByCategoria', [RutaCarreraController::class, 'getEspecialidadesByCategoria']);
    Route::post('/rutacarrera/getAreasByCategoria', [RutaCarreraController::class, 'getAreasByCategoria']);
    Route::post('/rutacarrera/getEspecialidadesByCategoriaCuerpo', [RutaCarreraController::class, 'getEspecialidadesByCategoriaCuerpo']);
    Route::post('/rutacarrera/getAreasByCategoriaEspecialidad', [RutaCarreraController::class, 'getAreasByCategoriaEspecialidad']);
    Route::post('/rutacarrera/getDetalleCargoRutaCarrera', [RutaCarreraController::class, 'getDetalleCargoRutaCarrera']);
    Route::get('/rutacarrera/getCuerposEspecialidadesAreasRutaCarrera', [RutaCarreraController::class, 'getCuerposEspecialidadesAreasRutaCarrera']);
    Route::get('/rutacarrera/getEspecialidadesRutas', [RutaCarreraController::class, 'getEspecialidadesRutas']);
    Route::get('/rutacarrera/getRutasFull', [RutaCarreraController::class, 'getRutasFull']);
    Route::get('/rutacarrera/getRutaCarreraActivos', [RutaCarreraController::class, 'getRutaCarreraActivos']);
    Route::post('/rutacarrera/getWidthByRutas', [RutaCarreraController::class, 'getWidthByRutas']);
    Route::get('/rutacarrera/export', [ExcelRutaCarreraController::class, 'export']);

    // Roles
    Route::post('/rol/getRoles', [RolController::class, 'getRoles']);
    Route::post('/rol/crearRoles', [RolController::class, 'crearRoles']);
    Route::post('/rol/actualizarRoles', [RolController::class, 'actualizarRoles']);
    Route::post('/rol/crearRolPrivilegios', [RolController::class, 'crearRolPrivilegios']);
    Route::post('/rol/actualizarRolPrivilegios', [RolController::class, 'actualizarRolPrivilegios']);
    Route::post('/rol/getRolPrivilegiosById', [RolController::class, 'getRolPrivilegiosById']);
    Route::post('/rol/eliminarRolPrivilegiosById', [RolController::class, 'eliminarRolPrivilegiosById']);
    Route::get('/rol/getRolesActivos', [RolController::class, 'getRolesActivos']);
    Route::get('/rol/getModulos', [RolController::class, 'getModulos']);

    // Usuarios
    Route::get('/usuario/getUsuariosFull', [UsuarioController::class, 'getUsuariosFull']);
    Route::post('/usuario/getUsuarios', [UsuarioController::class, 'getUsuarios']);
    Route::post('/usuario/crearUsuarios', [UsuarioController::class, 'crearUsuarios']);
    Route::post('/usuario/actualizarUsuarios', [UsuarioController::class, 'actualizarUsuarios']);
    Route::post('/usuario/crearUsuarioRol', [UsuarioController::class, 'crearUsuarioRol']);
    Route::post('/usuario/actualizarUsuarioRol', [UsuarioController::class, 'actualizarUsuarioRol']);
    Route::post('/usuario/getUsuariosRolesById', [UsuarioController::class, 'getUsuariosRolesById']);
    Route::post('/usuario/getPermisosByUser', [UsuarioController::class, 'getPermisosByUser']);
    Route::post('/usuario/eliminarUsuariosRolesId', [UsuarioController::class, 'eliminarUsuariosRolesId']);
    Route::get('/usuario/getRolPrivilegiosPantalla', [UsuarioController::class, 'getRolPrivilegiosPantalla']);
    Route::post('/usuario/getRolesByUsuarioId', [UsuarioController::class, 'getRolesByUsuarioId']);

    // Usuarios-Menu
    Route::post('/usuariomenu/crudAsignarMenus', [UsuarioMenuController::class, 'crudAsignarMenus']);

    // Escalafón de Cargos
    Route::post('escalafon/export', [ExcelEscalafonCargoController::class, 'export']);
    Route::post('escalafon/checkReforteEscalafon', [ExcelEscalafonCargoController::class, 'checkReforteEscalafon']);
});