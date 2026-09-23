from server import crc16_ibm
# Teltonika Codec 8 example payload should produce CRC 0x8612
hex_data = (
    "08010000016B40D8EA300100000000000000000000000000000001"
    "05021503010101425E0F01F10000601A014E000000000000000001"
)
print(f"CRC = {crc16_ibm(bytes.fromhex(hex_data)):04X}")
