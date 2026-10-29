from crewai import Agent, Crew, Process, Task
from crewai.agents.agent_builder.base_agent import BaseAgent
from crewai.project import CrewBase, agent, crew, task

from duty_swap_automation.models import ParsedEmail


@CrewBase
class SwapParserCrew:
    """Turns a free-text email into structured duty swap requests."""

    agents: list[BaseAgent]
    tasks: list[Task]

    agents_config = "config/agents.yaml"
    tasks_config = "config/tasks.yaml"

    @agent
    def swap_request_analyst(self) -> Agent:
        return Agent(config=self.agents_config["swap_request_analyst"], verbose=False)  # type: ignore[index]

    @task
    def extract_swap_request(self) -> Task:
        return Task(config=self.tasks_config["extract_swap_request"], output_pydantic=ParsedEmail)  # type: ignore[index]

    @crew
    def crew(self) -> Crew:
        return Crew(agents=self.agents, tasks=self.tasks, process=Process.sequential, verbose=False)
