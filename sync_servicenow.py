"""Entrada de tareas programadas: python -m sync_servicenow --mes AAAA-MM."""

import argparse
import json
from datetime import datetime
from zoneinfo import ZoneInfo

from app_logic import init_db, sincronizar_servicenow, exigir_permiso_actor
from core.permissions import ACTION_WRITE_CASES, ACTION_WRITE_INCIDENTS
from core.settings import config_value
from services.servicenow import ServiceNowError


def main():
    parser = argparse.ArgumentParser(description="Sincroniza ServiceNow con la base de la aplicación.")
    parser.add_argument("--tipo", choices=("casos", "incidentes"), default="casos")
    modo = parser.add_mutually_exclusive_group()
    modo.add_argument("--mes", help="AAAA-MM; predeterminado: mes actual en Colombia")
    modo.add_argument("--numero", help="Caso CS... o incidente INC...")
    parser.add_argument("--actor", default=config_value("SERVICENOW_SYNC_ACTOR"), help="Administrador activo local")
    args = parser.parse_args()
    try:
        exigir_permiso_actor(args.actor, ACTION_WRITE_CASES if args.tipo == "casos" else ACTION_WRITE_INCIDENTS)
        init_db()
        mes = args.mes or datetime.now(ZoneInfo("America/Bogota")).strftime("%Y-%m")
        result = sincronizar_servicenow(actor_email=args.actor, tipo=args.tipo,
                                      numero=args.numero, mes=None if args.numero else mes)
    except (ServiceNowError, ValueError, PermissionError) as error:
        parser.exit(1, f"{error}\n")
    except Exception:
        parser.exit(1, "No se pudo sincronizar. Revisa configuración y disponibilidad de la base.\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
