import asyncio

import click
from graph import Graph
from msgraph.generated.models.o_data_errors.o_data_error import ODataError


@click.group()
def cli():
    """Microsoft Graph API CLI for Outlook/Office 365"""
    pass


@cli.command()
def user():
    """Display current user information"""
    try:
        graph = Graph()
        user_info = asyncio.run(graph.get_user())
        if user_info:
            click.echo(f'Hello, {user_info.display_name}')
            # For Work/school accounts, email is in mail property
            # Personal accounts, email is in userPrincipalName
            click.echo(f'Email: {user_info.mail or user_info.user_principal_name}')
    except ODataError as e:
        click.echo(f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}", err=True)


@cli.command()
def token():
    """Display access token"""
    try:
        graph = Graph()
        access_token = asyncio.run(graph.get_user_token())
        click.echo(f"User token: {access_token}")
    except ODataError as e:
        click.echo(f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}", err=True)


@cli.command()
def inbox():
    """List inbox messages"""
    try:
        graph = Graph()
        message_page = asyncio.run(graph.get_inbox())
        if message_page and message_page.value:
            # Output each message's details
            for message in message_page.value:
                click.echo(f'Message: {message.subject}')
                if message.from_ and message.from_.email_address:
                    click.echo(f'  From: {message.from_.email_address.name or "NONE"}')
                else:
                    click.echo('  From: NONE')
                click.echo(f'  Status: {"Read" if message.is_read else "Unread"}')
                click.echo(f'  Received: {message.received_date_time}')
                click.echo()

            # If @odata.nextLink is present
            more_available = message_page.odata_next_link is not None
            click.echo(f'More messages available? {more_available}')
    except ODataError as e:
        click.echo(f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}", err=True)


if __name__ == "__main__":
    cli()
