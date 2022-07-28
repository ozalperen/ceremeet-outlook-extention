#Each instruction in this file creates a new layer
#Here we are getting our node as Base image
FROM ubuntu:latest
RUN mkdir -p /app
#setting working directory in the container
WORKDIR /app
RUN apt-get update
RUN apt-get -y install curl gnupg
RUN curl -sL https://deb.nodesource.com/setup_16.x  | bash -
RUN apt-get -y install nodejs
#Creating a new directory for app files and setting path in the container
#setting working directory in the container
#copying the package.json file(contains dependencies) from project source dir to container dir
COPY ./ /app
# installing the dependencies into the container
RUN npx office-addin-dev-certs install
#container exposed network port number
EXPOSE 3000
#command to run within the container
CMD ["npm", "start"]