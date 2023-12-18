# pip install speedtest-cli
import speedtest as st

# Set Best Server
server = st.Speedtest()
server.get_best_server()

# Test Download Speed
down = server.download()
down = round(down / 1000000, 3)  # Arredonda para 2 casas decimais
print(f"\nDownload Speed: {down} Mb/s")

# Test Upload Speed
up = server.upload()
up = round(up / 1000000, 3)  # Arredonda para 2 casas decimais
print(f"Upload Speed: {up} Mb/s")

# Test Ping
ping = server.results.ping
print(f"Ping Speed: {ping}  Mili")