import asyncio

from server import TCP_HOST, TCP_PORT, handle_tracker, init_db


async def main():
    init_db()
    server = await asyncio.start_server(handle_tracker, TCP_HOST, TCP_PORT)
    sockets = ", ".join(str(s.getsockname()) for s in server.sockets or [])
    print(f"[GPS] TCP listener running on {sockets}", flush=True)
    try:
        async with server:
            await server.serve_forever()
    except asyncio.CancelledError:
        pass
    finally:
        server.close()
        await server.wait_closed()
        print("[GPS] TCP listener stopped", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
