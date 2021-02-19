package zhang.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Collections;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ResultModel {
    private List<String> errorMessage = Collections.emptyList();
    private List<SpecInfo> specInfos = Collections.emptyList();
}
