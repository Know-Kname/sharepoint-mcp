"""SharePoint site tools — thin delegator to sub-modules."""

from mcp.server.fastmcp import FastMCP

from tools.cross_site_tools import register_cross_site_tools
from tools.provisioning_tools import register_provisioning_tools
from tools.read_tools import register_read_tools
from tools.write_tools import register_write_tools


def register_site_tools(mcp: FastMCP):
    """Register all SharePoint tools with the MCP server."""
    register_read_tools(mcp)
    register_write_tools(mcp)
    register_provisioning_tools(mcp)
    register_cross_site_tools(mcp)
