import logging

def run_plugin(context):
    """
    Demo Hybrid plugin entry point
    """
    logging.info("Demo Hybrid Plugin: Started execution.")
    job_name = context.get('jobFileName')
    logging.info(f"Demo Hybrid Plugin: Processing job '{job_name}'")

    # Simulate some work
    result = {
        "plugin_name": "DemoHybridPlugin",
        "processed_job": job_name,
        "status": "success"
    }

    logging.info(f"Demo Hybrid Plugin: Finished execution with result: {result}")
    return result

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    # Simple dummy execution for manual testing without argparse
    print("Running in Standalone Mode (Demo Hybrid Plugin)")
    context = {"jobFileName": "Standalone_CLI_Run"}
    run_plugin(context)
