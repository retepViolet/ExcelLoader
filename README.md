## Set up:
prisma db generate

uvicorn api:app --reload


## http requests:
POST: "/upload"

DELETE: "/delete"

POST: "/calculate"

GET: "/history"

GET: "/file_list"

See more about parameters in "test.ipynb" and "api.py"
