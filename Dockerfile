FROM syncthing/syncthing:latest

EXPOSE 3390
EXPOSE 22000/tcp
EXPOSE 21027/udp

VOLUME ["/var/syncthing"]

ENTRYPOINT ["/usr/bin/syncthing"]
CMD ["-home=/var/syn                                                                                   
