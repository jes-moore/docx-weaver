export IMAGE=docx-weaver

build-docker-image:
	docker build -t $(IMAGE) .
