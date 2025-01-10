-- Looking at Total Cases vs Population (US); Shows percentage of population that contracted Covid
select location, date, total_cases, population, (total_cases/population)*100 as pct_contacted
from cov_deaths
where location like '%States%' and continent is not null
order by 1,2

-- Looking at Total Deaths vs Total Cases (US);Likelihood of dying if you contract Covid in the US
select location, date, total_cases, total_deaths, (total_deaths/total_cases)*100 as pct_death
from cov_deaths
where location like '%States%' and continent is not null
order by 1,2

-- Look at countries with highest infection rate compared to population
select location, population, max(total_cases) as highest_inf_count, max((total_cases/population))*100 as pct_pop_inf
from cov_deaths
where continent is not null
group by location, population
order by pct_pop_inf desc

-- Showing countries with highest death count per population
select location, max(cast(total_deaths as int)) as tot_death_count
from cov_deaths
where continent is not null and total_deaths is not null
group by location
order by tot_death_count desc

-- Showing continent with highest death count 
select continent, max(cast(total_deaths as int)) as tot_death_count
from cov_deaths
where continent is not null and total_deaths is not null
group by continent
order by tot_death_count desc

-- Global Numbers 
select sum(new_cases) as tot_cases, sum(cast(new_deaths as int)) as tot_deaths, sum(cast(new_deaths as int))/sum(new_cases)*100 as pct_deaths
from cov_deaths
where continent is not null
order by 1,2

-- Looking at total population vs vaccination
WITH popvsvac (continent, location, date, population, new_vaccinations, roll_tot_vac) AS (
SELECT  dea.continent, dea.location, dea.date, dea.population, vac.new_vaccinations,
SUM(CAST(vac.new_vaccinations AS INT)) OVER (PARTITION BY dea.location ORDER BY dea.location, dea.date) AS roll_tot_vac
FROM cov_deaths dea
JOIN 
  cov_vacc vac
  ON dea.location = vac.location
  AND dea.date = vac.date
  WHERE 
    dea.continent IS NOT NULL
)
SELECT *, (roll_tot_vac/population)*100 as pct_vac
FROM popvsvac