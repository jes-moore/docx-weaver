FROM python:3.11.9-slim-bullseye
RUN apt-get update && apt-get install -y -qq --no-install-recommends git

# Create app-user/appuser home
RUN useradd -ms /bin/bash appuser
RUN mkdir /home/appuser/app
USER appuser

# Copy the app / install-reqs / extend python path / set wk
COPY --chown=appuser:appuser . /home/appuser/app
RUN pip install -r /home/appuser/app/requirements.txt
ENV PYTHONPATH "${PYTHONPATH}:/home/appuser/app/"
WORKDIR /home/appuser/app