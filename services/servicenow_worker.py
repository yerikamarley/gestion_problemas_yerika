"""Actualización periódica del servidor, sin credenciales ni sesión del navegador."""

from threading import Event, Thread

from core.servicenow_config import sincronizacion_habilitada
from core.settings import config_value
from services.servicenow import ServiceNowMFARequiredError


class CaseSyncWorker:
    def __init__(self, sincronizar):
        self._sincronizar = sincronizar
        self._wake = Event()
        self._stop = Event()
        self._requiere_login = False
        self._thread = Thread(target=self._run, name="servicenow-casos", daemon=True)
        self._thread.start()

    def despertar(self):
        self._requiere_login = False
        self._wake.set()

    def detener(self):
        self._stop.set()
        self._wake.set()
        self._thread.join(timeout=1)

    def ejecutar_ciclo(self):
        if not sincronizacion_habilitada() or self._requiere_login:
            return
        actor = config_value("SERVICENOW_SYNC_ACTOR")
        if actor:
            try:
                self._sincronizar(actor_email=actor)
            except ServiceNowMFARequiredError:
                self._requiere_login = True
            except Exception:
                # La fachada guarda el estado seguro en BD; el siguiente ciclo reintenta.
                pass

    def _run(self):
        while not self._stop.is_set():
            self._wake.clear()
            self.ejecutar_ciclo()
            try:
                intervalo = max(30, int(config_value("SERVICENOW_SYNC_INTERVAL_SECONDS", "120")))
            except ValueError:
                intervalo = 120
            self._wake.wait(intervalo)
