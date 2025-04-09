from mcp.server.fastmcp import FastMCP

# Create the MCP server
mcp = FastMCP("MCP Test")

@mcp.tool()
def echo(text: str) -> str:
    """Echo the provided text."""
    return f"You said: {text}"

if __name__ == "__main__":
    print("Starting MCP Test Server...")
    mcp.run()