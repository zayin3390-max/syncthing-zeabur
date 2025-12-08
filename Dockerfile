FROM alpine:3.19

# Install curl & tar
RUN apk add --no-cache curl tar

# Install Syncthing
RUN curl -L https://github.com/syncthing/syncthing/releases/download/v1.27.2/syncthing-linux-amd64-v1.27.2.tar.gz \
    | tar xz -C /usr/local/bin --strip-components=1 syncthing-linux-amd64-v1.27.2/syncthing

# Create Syncthing data directory
RUN mkdir -p /var/syncthing

EXPOSE 3390
EXPOSE 22000/tcp
EXPOSE 21027/udp

CMD ["/usr/local/bin/syncthing", "-home=/var/syncthing", "-no-browser", "-gui-address=0.0.0.0:${PORT}"]

