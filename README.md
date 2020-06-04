# RemoteAgentHealthSCOMServerReport
SCOM - Remote Agent Health Report Script

The script gets all your agents that have critical Health Service Watcher object, and loops through each one, checking to see:

* Is the server in maintenance mode?
* When was the server last communicating or reset?
* What are the management server assignments?
* Can we resolve the agent from DNS?
* Can we ping the agent now?
* Can we connect to the remote Service Control Manager?
* Can we get the status of Healthservice?
* If stopped, start it
* If disabled, fix it
* If someone uninstalled the agent, lets us know

This is really helpful when you have a large environment, and a large number of agents that are not communicating.
