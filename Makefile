export IMAGE=word-weaver

build-docker-image:
	docker build -t $(IMAGE) .
