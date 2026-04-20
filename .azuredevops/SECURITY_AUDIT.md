# AUDIT SÉCURITÉ - INFRA + ONBOARDING

## 📋 ÉTAT ACTUEL

### Architecture globale
```
INFRA-SHARED (pipeline-infra-shared.yml):
├─ Stage 8: Créer réseau partagé (TOUJOURS exécuté)
│   └─ stage-create-network.yml → VNet + subnet PE
├─ Stage 9: Durcissement réseau (OPTIONNEL si allowSecurity=true)
│   └─ stage-apply-network-security.yml → Private Endpoints + fermeture publique
└─ Stage 10: Résumé

ONBOARDING (pipeline-onboarding-companion.yml):
├─ Stage 6: Déployer Container App
├─ Stage 8: Sécurité compagnon (OPTIONNEL si allowSecurity=true)
│   └─ stage-apply-companion-security.yml → Vérifier isolement ressources
└─ Stage 9: Résumé
```

---

## ✅ POINTS POSITIFS

### 1. **Séparation nette infra vs compagnon**
- Infra crée le VNet/subnet UNE FOIS (toujours exécuté)
- Infra applique PE/fermeture publique côté partagé (optionnel)
- Compagnon valide sa sécurité SANS modifier infra

### 2. **Pattern cohérent et consistant**
- Même paramètre `allowSecurity` dans les deux pipelines
- Même logic: stage optionnel avant résumé final
- Dépendances explicites dans YAML

### 3. **Idempotence robuste**
- VNet/subnet: check + update (pas de crash si existe)
- All PE operations: `ensure_*` functions (best practice)
- Companion security: validation only (pas de création)

### 4. **Nommage standardisé et prévisible**
```
VNet: vnet-companion-shared-<env>
Subnet: snet-private-endpoints
CIDR: 10.90.0.0/16 & 10.90.10.0/24 (internes, pas en paramètres)
Companion blob: <companionName> (simple, isolé)
```

### 5. **Audit trail détaillé**
- Echo statements clairs pour chaque vérification
- Rapport final conforme/non-conforme
- Tags appliqués pour tracking

---

## ❌ LACUNES CRITIQUES

### 🔴 **1. DÉPENDANCES INCOHÉRENTES - BLOCKER**

**Problème:**
```yaml
# ApplyCompanionSecurity (onboarding stage 8)
dependsOn:
  - DeployContainerApp  ← Uniquement ça

# Mais si infra allowSecurity=true, il faut AUSSI:
#  - Vérifier que ApplyNetworkSecurity (infra stage 9) a tourné
#  - Sinon compagnon valide isolation sur ressources SANS PE
```

**Impact:** Si infra n'a PAS durcis le réseau mais compagnon oui → validation faussement positive

**Correction souhaitée:**
L'onboarding ne PEUT PAS savoir l'état infra après déploiement. Il faut soit:
- Option A: Stage ApplyCompanionSecurity valide que les ressources partagées ONT les PE
- Option B: Ajouter paramètre `sharedSecurityApplied` en input du compagnon

---

### 🔴 **2. VALIDATION VNET/SUBNET MANQUANTE**

**stage-apply-companion-security.yml ligne 42+:**
```bash
COMPANION deployment assume le VNet existe mais NE LE VALIDE PAS
# On ignore le VNet complètement !
```

**Impact:** Si infra a échoué à créer réseau → compagnon dit "✅ ok" pour des choses qu'il ne peut pas utiliser

**À ajouter:**
```bash
# Vérifier que le réseau partagé est bien accessible
VNET_NAME="vnet-companion-shared-${{ parameters.environment }}"
if ! az network vnet show --resource-group "$RG" --name "$VNET_NAME" >/dev/null 2>&1; then
  echo "❌ VNet $VNET_NAME INTROUVABLE - réseau partagé doit être créé d'abord"
  exit 1
fi
```

---

### 🔴 **3. ZERO VALIDATION RBAC RÉELLE**

**stage-apply-companion-security.yml ligne 125:**
```bash
echo "✅ MSI du compagnon assignée à Container App"
echo "✅ Accès Storage/AI Search limité via RBAC"
```

**Problème:** C'est du TEXT COSMÉTIQUE, pas une vraie validation

**À corriger:**
```bash
# Vérifier VRAIMENT les roles RBAC
MSI_PRINCIPAL_ID=$(az identity show --name "$MSI_NAME" --resource-group "$RG" --query principalId -o tsv)

# Storage roles:
STORAGE_ROLES=$(az role assignment list \
  --assignee "$MSI_PRINCIPAL_ID" \
  --scope "$(az storage account show --name "$STORAGE_NAME" --query id -o tsv)" \
  --query "length(@)" -o tsv)

if [ "$STORAGE_ROLES" -gt 0 ]; then
  echo "✅ MSI a $STORAGE_ROLES role(s) sur Storage"
else
  echo "❌ MSI AUCUN ROLE sur Storage - RBAC manquant!"
  exit 1
fi
```

---

### 🟡 **4. ISOLATION INTER-COMPAGNONS ABSENTE**

**Problème:** Deux compagnons dans le même Container Apps Environment partagé
- Pas de NetworkPolicy
- Pas de NSG rules
- Pas d'isolation de trafic

**Impact:** Compagnon A peut théoriquement communiquer avec Compagnon B

**À ajouter:** Network Policy ou NSG si nécessaire

---

### 🟡 **5. AI SEARCH - CONTRÔLE INCOMPLET**

**Validations manquantes:**
- ✅ Index existe
- ❌ Datasource du compagnon existe
- ❌ Indexer du compagnon existe
- ❌ Indexer a les bonnes permissions

```bash
# À ajouter après index check:
delete_ai_search_asset() {
  local asset_type="$1"
  local asset_name="$2"
  HTTP_STATUS=$(curl -s -o /dev/null -w "%{http_code}" \
    -X GET "$SEARCH_URL/$asset_type/$asset_name?api-version=2024-07-01" \
    -H "api-key: $SEARCH_KEY")
  
  if [ "$HTTP_STATUS" = "200" ]; then
    echo "✅ AI Search $asset_type/$asset_name présent"
  else
    echo "❌ AI Search $asset_type/$asset_name INTROUVABLE (HTTP $HTTP_STATUS)"
  fi
}

delete_ai_search_asset "datasources" "${{ parameters.aiSearchDataSourceName }}"
delete_ai_search_asset "indexers" "${{ parameters.aiSearchIndexerName }}"
```

---

### 🟡 **6. CONTAINER APP CONFIG PAS VALIDÉE**

**Contrôles manquants:**
- ✅ Container App existe
- ❌ Container App est sur le bon VNet/subnet
- ❌ Environment variables contiennent AppInsights connection
- ❌ MSI est bien assignée

```bash
# Ajouter validation:
CA_MSI=$(az containerapp show --name "$CONTAINER_APP_NAME" \
  --resource-group "$RG" \
  --query "identity.userAssignedIdentities" -o json)

if echo "$CA_MSI" | grep -q "\"$MSI_ID\""; then
  echo "✅ Container App a la bonne MSI"
else
  echo "❌ Container App MSI MAUVAISE ou MANQUANTE"
  exit 1
fi
```

---

## 🎯 VERDICT GLOBAL

| Aspect | Score | Justification |
|--------|-------|---|
| **Architecture & Pattern** | 9/10 | Excellente séparation, dépendances claires (sauf 1 issue) |
| **Code quality** | 7/10 | Idempotent, lisible, mais validations cosmétiques |
| **Production readiness** | 5/10 | Lacunes critiques sur dépendances et RBAC |
| **Security posture** | 6/10 | Vérifications structurelles oui, isolation logique manque |

---

## 📋 CHECKLIST CORRECTIONS

### CRITIQUE (avant production)
- [ ] Dépendances: compagnon doit valider que infra a appliqué sécurité
- [ ] Valider VNet/subnet existent avant toute opération
- [ ] Vérifications RBAC réelles (pas cosmetique)
- [ ] Valider datasource + indexer AI Search complet

### IMPORTANT
- [ ] Valider MSI est assignée à Container App
- [ ] Vérifier environment variables AppInsights/Storage
- [ ] Valider connexion au bon VNet

### NICE-TO-HAVE
- [ ] NetworkPolicy inter-compagnons
- [ ] Audit trail système
- [ ] Conformité baselines CIS
