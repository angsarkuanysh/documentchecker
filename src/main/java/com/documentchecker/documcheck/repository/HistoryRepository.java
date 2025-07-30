package com.documentchecker.documcheck.repository;

import com.documentchecker.documcheck.model.History;
import org.springframework.data.jpa.repository.JpaRepository;
import java.util.List;

public interface HistoryRepository extends JpaRepository<History, Long> {
    List<History> findByUserIdOrderByDateTimeDesc(Long userId);
}
