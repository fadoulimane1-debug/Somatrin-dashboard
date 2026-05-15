from django.db import models


class TechnicianEmployee(models.Model):
    """Techniciens de maintenance."""

    odoo_id = models.IntegerField(unique=True)
    name = models.CharField(max_length=255)
    email = models.EmailField(blank=True)
    department = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class MaintenanceIntervention(models.Model):
    """Modèle d'intervention de maintenance."""

    odoo_id = models.IntegerField(unique=True)
    name = models.CharField(max_length=255)
    metric = models.CharField(max_length=255)
    equipment_team = models.CharField(max_length=255, blank=True)
    responsible = models.CharField(max_length=255, blank=True)
    duration = models.CharField(max_length=50, blank=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class MaintenanceRequest(models.Model):
    """Demandes de maintenance (cache local optionnel)."""

    STATES = [
        ("new", "Nouvelle demande"),
        ("in_progress", "En cours"),
        ("done", "Terminé"),
        ("to_rectify", "A rectifier les relevés KM"),
    ]
    TYPES = [
        ("corrective", "Corrective"),
        ("preventive", "Préventive"),
        ("other", "Autre"),
    ]

    odoo_id = models.IntegerField(unique=True)
    name = models.CharField(max_length=255)
    equipment = models.CharField(max_length=255, blank=True)
    intervention_type = models.CharField(max_length=255, blank=True)
    category = models.CharField(max_length=255, blank=True)
    subcategory = models.CharField(max_length=255, blank=True)
    date_created = models.DateTimeField(null=True, blank=True)
    date_breakdown = models.DateTimeField(null=True, blank=True)
    date_repair = models.DateTimeField(null=True, blank=True)
    date_planned = models.DateTimeField(null=True, blank=True)
    date_close = models.DateTimeField(null=True, blank=True)
    maintenance_type = models.CharField(max_length=50, choices=TYPES, default="corrective")
    state = models.CharField(max_length=50, choices=STATES, default="new")
    team = models.CharField(max_length=255, blank=True)
    workshop = models.CharField(max_length=255, blank=True)
    responsible = models.CharField(max_length=255, blank=True)
    driver = models.CharField(max_length=255, blank=True)
    priority = models.IntegerField(default=0)
    estimated_time = models.CharField(max_length=50, blank=True)
    repair_time = models.CharField(max_length=50, blank=True)
    gap_time = models.CharField(max_length=50, blank=True)
    company = models.CharField(max_length=255, blank=True)
    maintenance_plan = models.CharField(max_length=255, blank=True)
    intervention_model = models.CharField(max_length=255, blank=True)
    unit_type = models.CharField(max_length=50, blank=True)
    udm = models.CharField(max_length=50, blank=True)
    last_reading = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    technicians = models.ManyToManyField(TechnicianEmployee, blank=True)
    is_done = models.BooleanField(default=False)
    sync_date = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-sync_date", "-id"]

    def __str__(self):
        return f"{self.name} ({self.state})"
