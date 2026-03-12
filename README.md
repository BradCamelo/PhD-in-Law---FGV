# Data Infrastructure for Experimental Research in Administrative Law

## PhD Research — FGV Direito SP

This repository contains the data infrastructure, scripts, and documentation associated with my doctoral research at FGV Direito SP. The dissertation investigates how experimental methods—particularly Randomized Controlled Trials (RCTs)—can be applied to empirical research in law.

While the thesis has a primarily methodological focus, examining the role of causal inference and experimental design in legal scholarship, it also includes the implementation of an original field experiment involving Brazilian municipalities. The empirical component evaluates institutional innovations in public procurement systems.

This repository serves as the central environment for organizing, processing, documenting, and sharing the data and code used throughout the research.

All publicly available datasets used in the project, as well as the scripts responsible for data collection, processing, and analysis, are made available here to ensure transparency, reproducibility, and methodological clarity.

Parts of the data infrastructure and analytical scripts contained in this repository were developed with the assistance of AI tools. In particular, several components of the codebase were elaborated in collaboration with Claude. All scripts were subsequently reviewed, adapted, and validated within the context of the research design.

## Research Overview

The dissertation investigates how experimental methods can contribute to the identification of causal effects in legal and institutional contexts, particularly in the domain of public procurement and administrative governance.

Traditional legal scholarship frequently relies on doctrinal analysis and qualitative interpretation. While these approaches remain essential, they often face limitations when addressing questions of policy effectiveness and institutional performance.

The thesis therefore pursues two complementary objectives:

1. Methodological objective
To examine the theoretical and methodological foundations of experimental research in law, with particular emphasis on causal inference and Randomized Controlled Trials.

2. Applied objective
To implement a field experiment involving Brazilian municipalities, evaluating institutional innovations in public procurement systems.

The empirical component serves as a demonstration of how experimental designs can be operationalized in legal and institutional research, contributing to the broader development of empirical legal studies.

## Purpose of this Repository

This repository has three main purposes.

1. Data Infrastructure

It provides the computational infrastructure supporting the empirical components of the research. The scripts included here retrieve and structure datasets from official public sources, generating municipal-level datasets used in the analysis.

2. Replication Environment

It functions as a replication environment, where the datasets used in the dissertation are stored, documented, and organized. The goal is to allow other researchers to reproduce the datasets used in the empirical analysis.

3. Research Code Repository

It provides access to the code used throughout the research process, including:

- Data collection 

- Data cleaning and transformation procedures

- Dataset construction 

- Matching and randomization algorithms

- Exploratory analysis scripts

Together, these components ensure that the empirical stages of the research remain transparent, auditable, and reproducible.

## Project Structure

The repository is organized into the following main components:

data/
    Raw and processed datasets used in the research

scripts/
    Python scripts used for data collection, cleaning, and dataset construction

outputs/
    Generated datasets and analytical outputs

docs/
    Documentation describing datasets, variables, and sources

Scripts contained in the scripts/ directory allow researchers to reconstruct the datasets used in the project directly from the original public sources whenever possible.

## Data Transparency

Transparency and reproducibility are central principles of the empirical research conducted in this project.

Whenever permitted by the terms of the original data providers, the datasets used in the research are stored directly in this repository. When direct storage is not possible, scripts are provided to automatically retrieve the data from the original public sources.

The repository therefore includes:

- Scripts for retrieving data from public APIs and official databases

- Clean datasets used in empirical analysis

- Documentation describing the origin and structure of each dataset

- Code used for data processing and integration

This approach ensures that the empirical components of the research remain fully transparent and replicable.

## Data Sources

The datasets used in this research rely primarily on official Brazilian public data sources, including:

IBGE — Instituto Brasileiro de Geografia e Estatística
Demographic and territorial data, including population and municipal characteristics.

SIDRA / IBGE APIs
Statistical indicators derived from census and survey data.

SAGRES — Sistema de Acompanhamento da Gestão dos Recursos da Sociedade (TCE-PB)
Administrative and fiscal information reported by municipalities to the Tribunal de Contas do Estado da Paraíba, including budget execution and financial management data.

TRAMITA — Sistema de Tramitação de Processos (TCE-PB)
Administrative and procedural data related to oversight activities, procurement processes, and accountability procedures monitored by the Tribunal de Contas do Estado da Paraíba.

Brazilian National Treasury — SICONFI / FINBRA
Municipal fiscal data, including revenue and expenditure indicators.

Atlas do Desenvolvimento Humano no Brasil — PNUD / IPEA
Human development indicators at the municipal level.

These sources provide the empirical basis for constructing structured datasets at the municipal level used in the experimental analysis.

## Reproducibility

Ensuring replicability is a central objective of this repository.

All datasets used in the dissertation are generated through documented scripts that rely on publicly available data sources.

Researchers interested in replicating the datasets can do so by running the scripts provided in the scripts/ directory. These scripts automatically retrieve the relevant data, process the raw datasets, and generate consolidated municipal datasets used in the empirical analysis.

By providing both the data infrastructure and the computational procedures, this repository enables other researchers to reproduce the empirical components of the research.

## Citation

If you use materials from this repository in academic work, please cite the associated doctoral research project:

Camelo, Bradson T. L.
Experimental Methods in Administrative Law: Causal Inference and Institutional Innovation in Public Procurement.
PhD Research — FGV Direito SP.

## License

This repository is intended exclusively for academic and research purposes.

All datasets included in this repository originate from official public databases maintained by Brazilian governmental institutions. The data used in this research consist exclusively of aggregated municipal-level information, such as demographic indicators, fiscal statistics, and development indexes. No personal or individually identifiable information is collected, processed, or stored in this repository.

The use and dissemination of these datasets are fully consistent with the Brazilian legal framework governing transparency and public information.

In particular, the data used in this project comply with:

- Brazilian Federal Constitution (1988), especially the constitutional principles of publicity and transparency in public administration (Article 37), as well as the right of access to public information.

- Brazilian General Data Protection Law (Lei Geral de Proteção de Dados — LGPD, Law No. 13.709/2018), since the repository does not process personal data, but only aggregated and publicly available institutional and statistical information.

All data included in this repository are therefore public and open data, originally made available by Brazilian public institutions for purposes of transparency, accountability, and research.

Whenever possible, the original data sources are preserved and referenced. In cases where redistribution is restricted by the original data provider, scripts are provided to retrieve the data directly from the official public sources.

Users of this repository remain responsible for complying with the licensing terms and citation requirements established by the original data providers.