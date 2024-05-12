FROM python:3.11.9-slim-bullseye
RUN apt-get update && apt-get install -y -qq --no-install-recommends git

# Create app-user/appuser home
RUN useradd -ms /bin/bash appuser
RUN mkdir /home/appuser/app
USER appuser

# Install Requirements
RUN pip install pip==24.0
RUN pip install bayoo-docx==0.2.14
RUN pip install pydantic-settings==2.2.1
RUN pip install tqdm==4.66.4
RUN pip install ipykernel==6.29.4
RUN pip install openai==1.28.1
RUN pip install pydantic==2.7.1
RUN pip install pandas==2.2.2

# Copy the app / extend python path / set wk
COPY --chown=appuser:appuser . /home/appuser/app
ENV PYTHONPATH "${PYTHONPATH}:/home/appuser/app/"
WORKDIR /home/appuser/app