"""Cross-site SharePoint tools: enumerate sites, list drives, copy/move files across sites/libraries."""

import json
import logging
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from tools._tool_helpers import _check_auth
from utils.graph_client import GraphClient

logger = logging.getLogger("sharepoint_cross_site")


def register_cross_site_tools(mcp: FastMCP):
    """Register tools that work across multiple SharePoint sites + document libraries."""

    @mcp.tool()
    async def list_all_sites(ctx: Context, search_query: str = "*", limit: int = 50) -> str:
        """List every SharePoint site the app can access in the tenant.

        Args:
            search_query: Graph search filter. Use "*" for all sites, or a name fragment.
            limit: Max sites to return (default 50, max 999).

        Returns: JSON array of {displayName, webUrl, id, description}.
        """
        logger.info(f"Tool called: list_all_sites query={search_query} limit={limit}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            gc = GraphClient(sp_ctx)

            endpoint = f"/sites?search={quote(search_query)}&$top={min(limit, 999)}"
            data = await gc.get(endpoint)
            sites = data.get("value", [])
            out = [
                {
                    "displayName": s.get("displayName") or s.get("name", "Unknown"),
                    "webUrl": s.get("webUrl"),
                    "id": s.get("id"),
                    "description": s.get("description", ""),
                }
                for s in sites
            ]
            return json.dumps({"count": len(out), "sites": out}, indent=2)
        except Exception as e:
            logger.error(f"list_all_sites error: {e}")
            raise

    @mcp.tool()
    async def list_drives_in_site(ctx: Context, site_id: str) -> str:
        """List all document libraries (drives) in a specific site.

        Args:
            site_id: Full Graph site ID (e.g. "tenant.sharepoint.com,siteCollGuid,siteGuid").
                Get via list_all_sites or get_site_info.

        Returns: JSON array of {name, id, webUrl, driveType, description}.
        """
        logger.info(f"Tool called: list_drives_in_site site_id={site_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            gc = GraphClient(sp_ctx)

            data = await gc.get(f"/sites/{site_id}/drives")
            drives = data.get("value", [])
            out = [
                {
                    "name": d.get("name"),
                    "id": d.get("id"),
                    "webUrl": d.get("webUrl"),
                    "driveType": d.get("driveType"),
                    "description": d.get("description", ""),
                }
                for d in drives
            ]
            return json.dumps({"count": len(out), "drives": out}, indent=2)
        except Exception as e:
            logger.error(f"list_drives_in_site error: {e}")
            raise

    @mcp.tool()
    async def copy_item(
        ctx: Context,
        source_drive_id: str,
        source_item_id: str,
        target_drive_id: str,
        target_parent_id: str,
        new_name: str = "",
    ) -> str:
        """Copy a file or folder from one drive to another (same or different site).

        Works cross-site and cross-library. Graph API performs server-side copy.
        Returns immediately with 202 Accepted + a monitor URL. The copy completes async.

        Args:
            source_drive_id: Drive ID containing the source item (from list_drives_in_site).
            source_item_id: Item ID to copy (from list_folder_contents or get_item_metadata).
            target_drive_id: Destination drive ID.
            target_parent_id: Destination folder ID. Use "root" for drive root.
            new_name: Optional new name at destination. Empty = keep original name.

        Returns: JSON with status, monitor_url for polling completion, target location.
        """
        logger.info(
            f"Tool called: copy_item src=({source_drive_id}/{source_item_id}) "
            f"-> dst=({target_drive_id}/{target_parent_id}) name={new_name!r}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            gc = GraphClient(sp_ctx)

            body = {
                "parentReference": {
                    "driveId": target_drive_id,
                    "id": target_parent_id,
                }
            }
            if new_name:
                body["name"] = new_name

            endpoint = f"/drives/{source_drive_id}/items/{source_item_id}/copy"
            result = await gc.post(endpoint, body)
            return json.dumps(
                {
                    "status": "accepted",
                    "note": "Server-side copy queued. Poll monitor_url for completion.",
                    "monitor_url": result.get("_monitor_url"),
                    "raw_response": result,
                },
                indent=2,
            )
        except Exception as e:
            logger.error(f"copy_item error: {e}")
            raise

    @mcp.tool()
    async def move_item(
        ctx: Context,
        source_drive_id: str,
        source_item_id: str,
        target_drive_id: str,
        target_parent_id: str,
        new_name: str = "",
    ) -> str:
        """Move a file or folder to a different drive/folder (same or different site).

        Unlike copy_item, move is synchronous and uses PATCH on the item with a new parentReference.
        Note: cross-tenant moves are not supported. Cross-site within same tenant: yes.

        Args:
            source_drive_id: Drive ID of source item.
            source_item_id: Item ID to move.
            target_drive_id: Destination drive ID.
            target_parent_id: Destination folder ID. Use "root" for drive root.
            new_name: Optional rename. Empty = keep original name.

        Returns: JSON with new item metadata (id, name, webUrl, parentReference).
        """
        logger.info(
            f"Tool called: move_item src=({source_drive_id}/{source_item_id}) "
            f"-> dst=({target_drive_id}/{target_parent_id}) name={new_name!r}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            gc = GraphClient(sp_ctx)

            body = {
                "parentReference": {
                    "driveId": target_drive_id,
                    "id": target_parent_id,
                }
            }
            if new_name:
                body["name"] = new_name

            endpoint = f"/drives/{source_drive_id}/items/{source_item_id}"
            result = await gc.patch(endpoint, body)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"move_item error: {e}")
            raise
