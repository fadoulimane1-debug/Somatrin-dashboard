from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import path, include

from core.views import redirect_by_role, RoleAwareLoginView

urlpatterns = [
    path('admin/', admin.site.urls),

    # Authentification
    path('login/', RoleAwareLoginView.as_view(), name='login'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('redirect/', redirect_by_role, name='redirect_by_role'),

    # Application
    path('', include('reporting.urls')),
]
