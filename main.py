import asyncio
import sys
from command import Command

async def main():
    cmd = Command()
    
    # Check if command line arguments are provided
    if len(sys.argv) > 1:
        # Join all arguments after the script name as the command
        command_input = ' '.join(sys.argv[1:])
        print(f"Running command: {command_input}")
        
        # Execute the command directly
        await cmd.execute_command(command_input)
    else:
        # Interactive mode
        cmd.welcome()
        try:
            await cmd.command_loop()
        except KeyboardInterrupt:
            print("\nExiting...")

if __name__ == "__main__":
    asyncio.run(main())