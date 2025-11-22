from fastapi import Request
from typing import Any


def get_db(request: Request) -> Any:
	"""Return the MongoDB database instance stored on the FastAPI app state.

	This avoids importing the `db` object from `main` (which would cause
	an import cycle) and allows route handlers to request the DB via
	`Depends(get_db)`.
	"""
	return request.app.state.db

