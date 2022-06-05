from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import re

hostName = "localhost"
serverPort = 8080
regions = ['Автозаводский р-н',
           'Приокский р-н',
           'Ленинский р-н']


class MyServer(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        if 'test' in self.path:
            id = int(re.findall('[0-9]+', self.path)[0])
            self.wfile.write(bytes('{"message": "Response from id=' + str(id) + '", '
                                   '"region": "' + regions[id] +'"}', "utf-8"))


if __name__ == "__main__":
    webServer = HTTPServer((hostName, serverPort), MyServer)
    print("Server started http://%s:%s" % (hostName, serverPort))

    try:
        webServer.serve_forever()
    except KeyboardInterrupt:
        pass

    webServer.server_close()
    print("Server stopped.")
