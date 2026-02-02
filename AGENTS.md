# Goal

Build a performant weights and biases clone using rust as a server and python as an API client. Choice of browser framework and language is up to the agents, but performance is the most important aspect. On charts, scrubbing and zooming should be instant. selecting, deselecting, and searching for runs should also be instant.
The clone should support connecting to a wandb server and calling wandb.init() on a run, then submitting per step metrics.
The client app should support all features of weights and biases, and performantly render multiple line charts per run.
It should work in a distributed training setup in a "baterries-included" manner and the python client should never crash the run or consume material resources on the same process.

Use puppeteer or other software to test browser features.
Include extensive testing.
