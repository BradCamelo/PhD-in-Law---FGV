# Data Infrastructure for Experimental Research in Administrative Law

## PhD Research — FGV Direito SP

This repository contains the data infrastructure, scripts, and documentation associated with my doctoral research at FGV Direito SP. The thesis has a primarily methodological focus, investigating how experimental methods—particularly Randomized Controlled Trials (RCTs)—can be applied to empirical research in law and public administration.

Although the dissertation is fundamentally concerned with the methodological foundations of causal inference in legal research, it also includes the implementation of an original field experiment involving Brazilian municipalities. This repository serves as the central environment for the organization, processing, and dissemination of the data and code used in the project.

All publicly available datasets used in the research, as well as the code responsible for data collection, processing, and analysis, will be made available here in order to ensure transparency, replicability, and methodological clarity.

## Research Overview

The dissertation investigates how experimental methods can contribute to the identification of causal effects in legal and institutional contexts, particularly in the domain of public procurement and administrative governance.

The central methodological concern of the thesis is to explore how legal scholars can move beyond descriptive or purely doctrinal approaches and adopt rigorous empirical strategies capable of identifying causal relationships. In this context, the research discusses the theoretical and methodological foundations of experimental designs in law, with particular emphasis on Randomized Controlled Trials.

To illustrate the practical application of these methods, the research includes the implementation of an RCT involving Brazilian municipalities, aimed at evaluating institutional innovations in public procurement systems. The experiment serves as a concrete case study demonstrating how experimental designs can be integrated into legal and institutional research.

## Purpose of this Repository

This repository has three main objectives.

First, it functions as the data infrastructure supporting the empirical components of the research. Scripts included here retrieve and structure datasets from official public sources, generating municipal-level datasets used in the empirical analysis.

Second, it serves as a replication environment, where all publicly accessible data used in the dissertation are stored and documented. The goal is to allow other researchers to reproduce the datasets used in the analysis.

Third, it provides access to the code used throughout the research process, including data collection pipelines, data cleaning procedures, dataset construction, and exploratory analysis.

## Data Transparency

All public datasets used in the research will be stored in this repository whenever permitted by the terms of the original data providers. When direct storage is not possible, scripts will be provided to automatically retrieve the data from the original public sources.

The repository therefore includes:

Scripts for retrieving data from public APIs and databases

Clean datasets used in empirical analysis

Documentation describing the origin and structure of each dataset

Code used for data processing and integration

The goal is to ensure that the empirical components of the research remain transparent and reproducible.

## Data Sources

The datasets used in this research rely primarily on official Brazilian public data sources, including:

IBGE (Instituto Brasileiro de Geografia e Estatística) — demographic and territorial data

SIDRA / IBGE APIs — census and statistical indicators

SAGRES — Sistema de Acompanhamento da Gestão dos Recursos da Sociedade (TCE-PB) — administrative and fiscal information reported by municipalities to the Tribunal de Contas do Estado da Paraíba, including budget execution and financial management data.

TRAMITA — Sistema de Tramitação de Processos (TCE-PB) — administrative and procedural data related to oversight activities, procurement processes, and accountability procedures monitored by the Tribunal de Contas do Estado da Paraíba.

Brazilian National Treasury (SICONFI / FINBRA) — municipal fiscal data

Atlas do Desenvolvimento Humano no Brasil (PNUD/IPEA) — human development indicators

These sources provide the empirical basis for constructing structured datasets at the municipal level.

## Reproducibility

Ensuring replicability of empirical legal research is a central goal of this repository. All datasets used in the dissertation are generated through documented scripts that rely on publicly available data sources.

Researchers interested in replicating the datasets can do so by running the scripts provided in the scripts/ directory.

The script retrieves data from the relevant public sources and produces consolidated municipal datasets.

## License

This repository is intended for academic and research purposes. Data remain subject to the licensing conditions established by their respective original providers.
